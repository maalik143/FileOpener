"""
╔══════════════════════════════════════════════════════════════╗
║    UNIVERSAL FILE OPENER · EXTRACTOR · CONVERTER  v2.0       ║
║    Tab 1: Extract — ZIP, RAR, 7Z, TAR, GZ …                 ║
║    Tab 2: Convert — IMG↔IMG, DOCX→PDF, CSV↔XLSX, TXT→PDF    ║
║    Tab 3: Batch  — Extract multiple archives at once         ║
║    Features: Drag&Drop · Progress · Large GB · Dark UI       ║
║             Windows Context Menu Integration                  ║
╚══════════════════════════════════════════════════════════════╝
"""
from __future__ import annotations
import os, sys, zipfile, tarfile, shutil, subprocess, threading, io, winreg
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime

# ── Drag & Drop support ──────────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ── Optional imports ─────────────────────────────────────────
try:    import py7zr;    HAS_7Z  = True
except: HAS_7Z  = False
try:    import rarfile;  HAS_RAR = True
except: HAS_RAR = False
try:
    from PIL import Image
    HAS_PIL = True
except: HAS_PIL = False
try:    import pandas as pd;     HAS_PD  = True
except: HAS_PD  = False
try:    import openpyxl;         HAS_XL  = True
except: HAS_XL  = False
try:    from docx import Document; HAS_DOCX = True
except: HAS_DOCX = False
try:    from fpdf import FPDF;   HAS_FPDF = True
except: HAS_FPDF = False
try:    from reportlab.pdfgen import canvas as rl_canvas; HAS_RL = True
except: HAS_RL = False

# ── Colours ──────────────────────────────────────────────────
BG        = "#0d1117"
SURFACE   = "#161b22"
SURFACE2  = "#21262d"
BORDER    = "#30363d"
ACCENT    = "#7c3aed"
ACCENT_LT = "#a78bfa"
GREEN     = "#10b981"
RED       = "#ef4444"
YELLOW    = "#f59e0b"
CYAN      = "#06b6d4"
PINK      = "#ec4899"
TEXT      = "#e6edf3"
TEXT2     = "#8b949e"

def fmt_size(b):
    for u in ("B","KB","MB","GB","TB"):
        if b < 1024: return f"{b:.1f} {u}"
        b /= 1024
    return f"{b:.1f} PB"

def get_type(path):
    p = path.lower()
    if p.endswith(".zip"):                       return "zip"
    if p.endswith(".rar"):                       return "rar"
    if p.endswith((".7z",".7zip")):              return "7z"
    if p.endswith((".tar.gz",".tgz")):           return "tar.gz"
    if p.endswith((".tar.bz2",".tbz2")):         return "tar.bz2"
    if p.endswith((".tar.xz",".txz")):           return "tar.xz"
    if p.endswith(".tar"):                       return "tar"
    if p.endswith(".gz"):                        return "gz"
    if p.endswith(".bz2"):                       return "bz2"
    if p.endswith(".xz"):                        return "xz"
    if p.endswith(".pdf"):                       return "pdf"
    if p.endswith((".png",".jpg",".jpeg",".gif",".bmp",".webp",".tiff",".tif",".ico")): return "image"
    if p.endswith((".mp4",".mkv",".avi",".mov",".wmv",".flv",".webm")):                 return "video"
    if p.endswith((".mp3",".wav",".flac",".ogg",".aac")):                               return "audio"
    if p.endswith((".docx",".doc")):             return "word"
    if p.endswith((".xlsx",".xls")):             return "excel"
    if p.endswith(".csv"):                       return "csv"
    if p.endswith((".pptx",".ppt")):             return "ppt"
    # ── Text & code types ─────────────────────────────────
    if p.endswith((".html", ".htm")):            return "html"
    if p.endswith((".py",".pyw")):               return "python"
    if p.endswith((".js",".ts",".jsx",".tsx")):  return "javascript"
    if p.endswith(".json"):                      return "json"
    if p.endswith(".xml"):                       return "xml"
    if p.endswith(".css"):                       return "css"
    if p.endswith((".txt",".log",".md",".rst",".ini",".cfg",".yaml",".yml",".toml")): return "text"
    if p.endswith((".c",".cpp",".h",".hpp",".cs",".java",".go",".rs",".php",".rb",".sh",".bat",".ps1",".swift",".kt")): return "code"
    if p.endswith(".sql"):                       return "sql"
    if p.endswith(".svg"):                       return "svg"
    return "other"

def is_archive(t):
    return t in {"zip","rar","7z","tar.gz","tar.bz2","tar.xz","tar","gz","bz2","xz"}

# ── Conversion map: {src_type: [(label, ext), …]} ────────────
CONVERSIONS = {
    "image": [
        ("PNG  →  JPG",   ".jpg"),
        ("PNG  →  BMP",   ".bmp"),
        ("PNG  →  WEBP",  ".webp"),
        ("PNG  →  TIFF",  ".tiff"),
        ("PNG  →  ICO",   ".ico"),
        ("JPG  →  PNG",   ".png"),
        ("JPG  →  WEBP",  ".webp"),
        ("JPG  →  BMP",   ".bmp"),
        ("BMP  →  PNG",   ".png"),
        ("BMP  →  JPG",   ".jpg"),
        ("WEBP →  PNG",   ".png"),
        ("WEBP →  JPG",   ".jpg"),
        ("GIF  →  PNG",   ".png"),
        ("TIFF →  PNG",   ".png"),
        ("TIFF →  JPG",   ".jpg"),
        ("ANY IMAGE → ICO", ".ico"),
        ("ANY IMAGE → PDF", ".pdf"),
    ],
    "word": [
        ("DOCX  →  PDF",  ".pdf"),
        ("DOCX  →  TXT",  ".txt"),
    ],
    "text": [
        ("TXT   →  PDF",  ".pdf"),
        ("TXT   →  HTML", ".html"),
        ("MD    →  TXT",  ".txt"),
        ("TXT   →  JSON", ".json"),
    ],
    "html": [
        ("HTML  →  PDF",  ".pdf"),
        ("HTML  →  TXT",  ".txt"),
        ("HTML  →  MD",   ".md"),
    ],
    "python": [
        ("PY    →  PDF",  ".pdf"),
        ("PY    →  TXT",  ".txt"),
        ("PY    →  HTML (highlight)", ".html"),
    ],
    "javascript": [
        ("JS/TS →  PDF",  ".pdf"),
        ("JS/TS →  TXT",  ".txt"),
    ],
    "json": [
        ("JSON  →  CSV",  ".csv"),
        ("JSON  →  TXT",  ".txt"),
        ("JSON  →  XLSX", ".xlsx"),
    ],
    "xml": [
        ("XML   →  TXT",  ".txt"),
        ("XML   →  JSON", ".json"),
    ],
    "css": [
        ("CSS   →  TXT",  ".txt"),
        ("CSS   →  PDF",  ".pdf"),
    ],
    "code": [
        ("CODE  →  PDF",  ".pdf"),
        ("CODE  →  TXT",  ".txt"),
        ("CODE  →  HTML", ".html"),
    ],
    "sql": [
        ("SQL   →  PDF",  ".pdf"),
        ("SQL   →  TXT",  ".txt"),
    ],
    "svg": [
        ("SVG   →  PNG",  ".png"),
        ("SVG   →  PDF",  ".pdf"),
    ],
    "csv": [
        ("CSV   →  XLSX", ".xlsx"),
        ("CSV   →  TXT",  ".txt"),
        ("CSV   →  JSON", ".json"),
    ],
    "excel": [
        ("XLSX  →  CSV",  ".csv"),
        ("XLSX  →  JSON", ".json"),
        ("XLSX  →  TXT",  ".txt"),
    ],
    "pdf": [
        ("PDF   →  TXT",  ".txt"),
        ("PDF   →  PNG (pages)", ".png"),
    ],
    "zip": [
        ("ZIP   →  7Z",   ".7z"),
        ("ZIP   →  TAR.GZ",".tar.gz"),
    ],
    "7z": [
        ("7Z    →  ZIP",  ".zip"),
    ],
}

TYPE_ICONS = {
    "zip":"📦","rar":"📦","7z":"📦","tar.gz":"📦","tar.bz2":"📦",
    "tar.xz":"📦","tar":"📦","gz":"📦","bz2":"📦","xz":"📦",
    "pdf":"📄","image":"🖼️","video":"🎬","audio":"🎵",
    "word":"📝","excel":"📊","csv":"📊","ppt":"📋","text":"📃",
    "html":"🌐","python":"🐍","javascript":"⚡","json":"{ }",
    "xml":"<>","css":"🎨","code":"💻","sql":"🗄️","svg":"🎭","other":"📁",
}
TYPE_COLORS = {
    "zip":ACCENT,"rar":ACCENT,"7z":ACCENT,"tar.gz":ACCENT,
    "tar.bz2":ACCENT,"tar.xz":ACCENT,"tar":ACCENT,"gz":ACCENT,
    "bz2":ACCENT,"xz":ACCENT,"pdf":RED,"image":GREEN,
    "video":CYAN,"audio":YELLOW,"word":"#3b82f6","excel":GREEN,
    "csv":"#22c55e","ppt":"#f97316","text":TEXT2,
    "html":CYAN,"python":YELLOW,"javascript":"#f0db4f","json":"#89e051",
    "xml":"#f97316","css":"#563d7c","code":ACCENT_LT,"sql":"#cc2927",
    "svg":PINK,"other":TEXT2,
}

# ══════════════════════════════════════════════════════════════
class FileOpener(TkinterDnD.Tk if HAS_DND else tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("Universal File Tool — Open · Extract · Convert  v2.0")
        self.geometry("920x700")
        self.minsize(750, 560)
        self.configure(bg=BG)
        self.resizable(True, True)

        self.current_file: str | None = None
        self.extract_dest: str | None = None
        self._cancel = False

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w,h = self.winfo_width(), self.winfo_height()
        sw,sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    # ── Helpers ───────────────────────────────────────────────
    def _lighten(self, c):
        try:
            r=int(c[1:3],16); g=int(c[3:5],16); b=int(c[5:7],16)
            return f"#{min(255,r+30):02x}{min(255,g+30):02x}{min(255,b+30):02x}"
        except: return c

    def _btn(self, parent, text, cmd, bg, fg,
             font=("Segoe UI",10,"bold"), padx=12, pady=6):
        b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                      font=font, relief="flat", bd=0,
                      activebackground=SURFACE2, activeforeground=TEXT,
                      cursor="hand2", padx=padx, pady=pady)
        b.bind("<Enter>", lambda e: b.configure(bg=self._lighten(bg)))
        b.bind("<Leave>", lambda e: b.configure(bg=bg))
        return b

    def _log(self, msg, tag="info"):
        ts = datetime.now().strftime("%H:%M:%S")
        for lw in (self.log1, self.log2):
            lw.configure(state="normal")
            lw.insert("end", f"[{ts}] {msg}\n", tag)
            lw.see("end")
            lw.configure(state="disabled")
        self.status_var.set(msg[:100])

    def _set_prog(self, pct, msg=""):
        self.progress_var.set(pct)
        self.prog_lbl.configure(text=msg)
        self.conv_progress_var.set(pct)
        self.conv_prog_lbl.configure(text=msg)

    # ── BUILD UI ─────────────────────────────────────────────
    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg=SURFACE, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="⚡ Universal File Tool",
                 font=("Segoe UI",17,"bold"), fg=TEXT, bg=SURFACE).pack(side="left", padx=18)
        tk.Label(hdr, text="Extract · Convert · Open — Any Format",
                 font=("Segoe UI",9), fg=TEXT2, bg=SURFACE).pack(side="left")

        # Notebook (tabs)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab",
                         background=SURFACE2, foreground=TEXT2,
                         padding=[18,8], font=("Segoe UI",10,"bold"))
        style.map("TNotebook.Tab",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "white")])
        style.configure("Custom.Horizontal.TProgressbar",
                         troughcolor=SURFACE2, background=ACCENT,
                         bordercolor=BG, lightcolor=ACCENT_LT, darkcolor=ACCENT)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=6)

        # ── TAB 1: EXTRACT ────────────────────────────────────
        tab1 = tk.Frame(self.nb, bg=BG)
        self.nb.add(tab1, text="📦  Furi / Extract")
        self._build_tab_extract(tab1)

        # ── TAB 2: CONVERT ───────────────────────────────────
        tab2 = tk.Frame(self.nb, bg=BG)
        self.nb.add(tab2, text="🔄  Badal / Convert")
        self._build_tab_convert(tab2)

        # ── TAB 3: BATCH EXTRACT ─────────────────────────────
        tab3 = tk.Frame(self.nb, bg=BG)
        self.nb.add(tab3, text="📂  Xog Badan / Batch")
        self._build_tab_batch(tab3)

        # ── TAB 4: SETTINGS ──────────────────────────────────
        tab4 = tk.Frame(self.nb, bg=BG)
        self.nb.add(tab4, text="⚙️  Dejinta / Settings")
        self._build_tab_settings(tab4)

        # Shared progress & log
        bottom = tk.Frame(self, bg=BG)
        bottom.pack(fill="both", expand=False, padx=10, pady=(0,6))

        # Progress
        pf = tk.Frame(bottom, bg=BG)
        pf.pack(fill="x")
        self.progress_var     = tk.DoubleVar()
        self.conv_progress_var = self.progress_var          # same var
        self.pbar = ttk.Progressbar(pf, variable=self.progress_var,
                                     maximum=100, style="Custom.Horizontal.TProgressbar")
        self.pbar.pack(fill="x", ipady=5)
        self.prog_lbl = tk.Label(pf, text="", font=("Segoe UI",9),
                                  fg=TEXT2, bg=BG)
        self.prog_lbl.pack(anchor="w")
        self.conv_prog_lbl = self.prog_lbl   # same widget

        # Log
        lf = tk.Frame(bottom, bg=BG)
        lf.pack(fill="both", expand=True, pady=(4,0))
        lhdr = tk.Frame(lf, bg=BG)
        lhdr.pack(fill="x")
        tk.Label(lhdr, text="📋 Log:", font=("Segoe UI",9,"bold"),
                 fg=TEXT2, bg=BG).pack(side="left")
        self._btn(lhdr, "Clear", self._clear_log, BG, TEXT2,
                  padx=8, pady=2).pack(side="right")

        li = tk.Frame(lf, bg=SURFACE2,
                       highlightbackground=BORDER, highlightthickness=1)
        li.pack(fill="both", expand=True)
        lsb = tk.Scrollbar(li); lsb.pack(side="right", fill="y")

        self.log1 = tk.Text(li, bg=SURFACE2, fg=TEXT,
                             font=("Consolas",9), relief="flat", bd=6,
                             state="disabled", height=5,
                             yscrollcommand=lsb.set)
        self.log1.pack(fill="both", expand=True)
        lsb.config(command=self.log1.yview)
        self.log2 = self.log1   # same widget, both tabs share
        for tag, col in [("ok",GREEN),("err",RED),("info",CYAN),("warn",YELLOW)]:
            self.log1.tag_configure(tag, foreground=col)

        # Status bar
        self.status_var = tk.StringVar(value="Diyaar — fayl soo dhig ama browse")
        tk.Label(self, textvariable=self.status_var,
                 font=("Segoe UI",9), fg=TEXT2, bg=SURFACE,
                 anchor="w", pady=4, padx=14).pack(fill="x", side="bottom")

    # ══════════════════════════════════════════════════════════
    # TAB 1 — EXTRACT
    # ══════════════════════════════════════════════════════════
    def _build_tab_extract(self, parent):
        # Drop zone — Real Drag & Drop
        dz = tk.Frame(parent, bg=SURFACE2,
                       highlightbackground=ACCENT, highlightthickness=2,
                       cursor="hand2")
        dz.pack(fill="x", padx=12, pady=(10,4))
        dz.bind("<Button-1>", lambda e: self._browse())

        self.drop_lbl = tk.Label(dz,
            text="🗂️   Jiid & Dhig (Drag & Drop) Archive Halkan   —   ama   Click Browse",
            font=("Segoe UI",13,"bold"), fg=ACCENT_LT, bg=SURFACE2, cursor="hand2")
        self.drop_lbl.pack(pady=(18,2))
        self.drop_lbl.bind("<Button-1>", lambda e: self._browse())
        dnd_status = "✅ Drag & Drop waa diyaar" if HAS_DND else "⚠️ Drag & Drop: pip install tkinterdnd2"
        tk.Label(dz, text=f"ZIP · RAR · 7Z · TAR · GZ · BZ2 · XZ    {dnd_status}",
                 font=("Segoe UI",9), fg=GREEN if HAS_DND else YELLOW, bg=SURFACE2).pack(pady=(0,16))

        # Enable real drag and drop
        if HAS_DND:
            dz.drop_target_register(DND_FILES)
            dz.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_lbl.drop_target_register(DND_FILES)
            self.drop_lbl.dnd_bind("<<Drop>>", self._on_drop)

        # File info
        info = tk.Frame(parent, bg=SURFACE,
                         highlightbackground=BORDER, highlightthickness=1)
        info.pack(fill="x", padx=12, pady=4)

        top = tk.Frame(info, bg=SURFACE)
        top.pack(fill="x", padx=12, pady=(8,4))
        self.type_lbl = tk.Label(top, text="", font=("Segoe UI",26),
                                  bg=SURFACE, fg=ACCENT)
        self.type_lbl.pack(side="left")
        nc = tk.Frame(top, bg=SURFACE); nc.pack(side="left", padx=8)
        self.name_lbl = tk.Label(nc, text="Fayl ma jiro weli",
                                  font=("Segoe UI",12,"bold"),
                                  fg=TEXT, bg=SURFACE, wraplength=550, justify="left")
        self.name_lbl.pack(anchor="w")
        self.meta_lbl = tk.Label(nc, text="",
                                  font=("Segoe UI",9), fg=TEXT2, bg=SURFACE)
        self.meta_lbl.pack(anchor="w")

        # Contents list
        cl = tk.Frame(info, bg=SURFACE); cl.pack(fill="both", expand=True, padx=12, pady=(0,8))
        self.contents_label = tk.Label(cl, text="",
                                        font=("Segoe UI",9,"bold"),
                                        fg=ACCENT_LT, bg=SURFACE)
        self.contents_label.pack(anchor="w")
        csb = tk.Scrollbar(cl); csb.pack(side="right", fill="y")
        self.contents_list = tk.Listbox(cl, bg=SURFACE2, fg=TEXT,
                                         selectbackground=ACCENT,
                                         font=("Consolas",9),
                                         relief="flat", bd=0, height=7,
                                         yscrollcommand=csb.set)
        self.contents_list.pack(fill="both", expand=True)
        csb.config(command=self.contents_list.yview)

        # Destination
        df = tk.Frame(parent, bg=BG, padx=12); df.pack(fill="x", pady=4)
        tk.Label(df, text="📂 Meesha la furi doono:", font=("Segoe UI",10,"bold"),
                 fg=TEXT2, bg=BG).pack(anchor="w")
        dr = tk.Frame(df, bg=BG); dr.pack(fill="x", pady=3)
        self.dest_var = tk.StringVar(value="(Isla folder-ka faylka)")
        tk.Entry(dr, textvariable=self.dest_var, font=("Segoe UI",10),
                 fg=TEXT, bg=SURFACE2, relief="flat", bd=6,
                 insertbackground=TEXT).pack(side="left", fill="x", expand=True, ipady=5)
        self._btn(dr,"📁 Dooro",self._choose_dest,SURFACE2,TEXT).pack(side="left",padx=4)
        self._btn(dr,"🏠 Default",self._reset_dest,SURFACE2,TEXT2).pack(side="left")

        # Buttons
        bf = tk.Frame(parent, bg=BG, padx=12, pady=8); bf.pack(fill="x")
        self._btn(bf,"📦  Furi / Extract",self._start_extract,ACCENT,"white",
                  font=("Segoe UI",12,"bold"),padx=22,pady=10).pack(side="left",padx=(0,6))
        self._btn(bf,"🚀  Fur Faylka",self._open_file,GREEN,"white",
                  font=("Segoe UI",12,"bold"),padx=22,pady=10).pack(side="left",padx=(0,6))
        self._btn(bf,"⛔  Cancel",self._cancel_extract,SURFACE2,RED,
                  font=("Segoe UI",11,"bold"),padx=14,pady=10).pack(side="left",padx=(0,6))
        self._btn(bf,"📂 Fur Folder",self._open_dest,SURFACE2,TEXT2,
                  padx=14,pady=10).pack(side="right",padx=(4,0))
        self._btn(bf,"🗑️ Clear",self._clear,SURFACE2,TEXT2,
                  padx=14,pady=10).pack(side="right",padx=4)

    # ══════════════════════════════════════════════════════════
    # TAB 2 — CONVERT
    # ══════════════════════════════════════════════════════════
    def _build_tab_convert(self, parent):
        # Input file
        inp = tk.Frame(parent, bg=SURFACE,
                        highlightbackground=BORDER, highlightthickness=1)
        inp.pack(fill="x", padx=12, pady=(10,4))

        tk.Label(inp, text="1️⃣  Dooro Faylka Aad Badali Rabto (Input):",
                 font=("Segoe UI",10,"bold"), fg=TEXT2, bg=SURFACE).pack(anchor="w", padx=12, pady=(8,4))
        ir = tk.Frame(inp, bg=SURFACE); ir.pack(fill="x", padx=12, pady=(0,8))
        self.conv_src_var = tk.StringVar(value="")
        self.conv_src_entry = tk.Entry(ir, textvariable=self.conv_src_var,
                                        font=("Segoe UI",10), fg=TEXT,
                                        bg=SURFACE2, relief="flat", bd=6,
                                        insertbackground=TEXT)
        self.conv_src_entry.pack(side="left", fill="x", expand=True, ipady=6)
        self._btn(ir,"📁 Browse",self._conv_browse,ACCENT,"white",
                  padx=14,pady=6).pack(side="left",padx=(6,0))

        # File type info badge
        self.conv_info_lbl = tk.Label(inp, text="",
                                       font=("Segoe UI",9), fg=TEXT2, bg=SURFACE)
        self.conv_info_lbl.pack(anchor="w", padx=14, pady=(0,8))

        # Conversion type selector
        sel = tk.Frame(parent, bg=SURFACE,
                        highlightbackground=BORDER, highlightthickness=1)
        sel.pack(fill="x", padx=12, pady=4)

        tk.Label(sel, text="2️⃣  Dooro Nooca Badalka (Conversion Type):",
                 font=("Segoe UI",10,"bold"), fg=TEXT2, bg=SURFACE).pack(anchor="w", padx=12, pady=(8,4))

        conv_row = tk.Frame(sel, bg=SURFACE); conv_row.pack(fill="x", padx=12, pady=(0,10))

        self.conv_type_var = tk.StringVar(value="— marka hore fayl dooro —")
        self.conv_dropdown = ttk.Combobox(conv_row, textvariable=self.conv_type_var,
                                           font=("Segoe UI",11), state="readonly",
                                           width=40)
        self.conv_dropdown.pack(side="left", fill="x", expand=True, ipady=4)

        # Style combobox
        style2 = ttk.Style()
        style2.configure("TCombobox",
                          fieldbackground=SURFACE2, background=SURFACE2,
                          foreground=TEXT, selectbackground=ACCENT)

        # Output destination
        out_frame = tk.Frame(parent, bg=SURFACE,
                              highlightbackground=BORDER, highlightthickness=1)
        out_frame.pack(fill="x", padx=12, pady=4)

        tk.Label(out_frame, text="3️⃣  Meesha la Kaydinayo (Output Folder):",
                 font=("Segoe UI",10,"bold"), fg=TEXT2, bg=SURFACE).pack(anchor="w", padx=12, pady=(8,4))
        outr = tk.Frame(out_frame, bg=SURFACE); outr.pack(fill="x", padx=12, pady=(0,8))
        self.conv_dest_var = tk.StringVar(value="(Isla folder-ka faylka — default)")
        tk.Entry(outr, textvariable=self.conv_dest_var, font=("Segoe UI",10),
                 fg=TEXT, bg=SURFACE2, relief="flat", bd=6,
                 insertbackground=TEXT).pack(side="left", fill="x", expand=True, ipady=5)
        self._btn(outr,"📁 Dooro",self._conv_choose_dest,SURFACE2,TEXT,
                  padx=12,pady=6).pack(side="left",padx=4)

        # QUALITY / OPTIONS row
        opt_frame = tk.Frame(parent, bg=SURFACE,
                              highlightbackground=BORDER, highlightthickness=1)
        opt_frame.pack(fill="x", padx=12, pady=4)

        tk.Label(opt_frame, text="4️⃣  Xulashooyinka (Options):",
                 font=("Segoe UI",10,"bold"), fg=TEXT2, bg=SURFACE).pack(anchor="w", padx=12, pady=(8,4))

        opts_row = tk.Frame(opt_frame, bg=SURFACE); opts_row.pack(fill="x", padx=12, pady=(0,10))

        # Image quality
        tk.Label(opts_row, text="Tayada Sawirka:", font=("Segoe UI",9),
                 fg=TEXT2, bg=SURFACE).pack(side="left")
        self.quality_var = tk.IntVar(value=92)
        tk.Scale(opts_row, from_=30, to=100, orient="horizontal",
                 variable=self.quality_var, bg=SURFACE, fg=TEXT,
                 troughcolor=SURFACE2, highlightthickness=0,
                 activebackground=ACCENT, length=150).pack(side="left", padx=6)
        tk.Label(opts_row, textvariable=tk.StringVar(), font=("Segoe UI",9),
                 fg=ACCENT_LT, bg=SURFACE).pack(side="left")

        # Image resize
        tk.Label(opts_row, text="  Resize (%):", font=("Segoe UI",9),
                 fg=TEXT2, bg=SURFACE).pack(side="left", padx=(12,0))
        self.resize_var = tk.IntVar(value=100)
        tk.Scale(opts_row, from_=10, to=200, orient="horizontal",
                 variable=self.resize_var, bg=SURFACE, fg=TEXT,
                 troughcolor=SURFACE2, highlightthickness=0,
                 activebackground=GREEN, length=150).pack(side="left", padx=6)

        # PDF paper size
        tk.Label(opts_row, text="  PDF Size:", font=("Segoe UI",9),
                 fg=TEXT2, bg=SURFACE).pack(side="left", padx=(12,0))
        self.pdf_size_var = tk.StringVar(value="A4")
        ttk.Combobox(opts_row, textvariable=self.pdf_size_var,
                     values=["A4","A3","Letter","Legal"],
                     state="readonly", width=8,
                     font=("Segoe UI",9)).pack(side="left", padx=4)

        # CONVERT BUTTON
        cbf = tk.Frame(parent, bg=BG, padx=12, pady=10); cbf.pack(fill="x")
        self._btn(cbf,"🔄  BADAL / CONVERT NOW",self._start_convert,PINK,"white",
                  font=("Segoe UI",13,"bold"),padx=30,pady=12).pack(side="left")
        self._btn(cbf,"📂 Fur Output Folder",self._conv_open_dest,GREEN,"white",
                  padx=14,pady=12).pack(side="left",padx=8)
        self._btn(cbf,"⛔ Cancel",self._cancel_extract,SURFACE2,RED,
                  padx=12,pady=12).pack(side="left")

        # Preview area
        prev = tk.Frame(parent, bg=SURFACE,
                         highlightbackground=BORDER, highlightthickness=1)
        prev.pack(fill="both", expand=True, padx=12, pady=4)
        tk.Label(prev, text="👁️ Preview / Natiijada:",
                 font=("Segoe UI",9,"bold"), fg=TEXT2, bg=SURFACE).pack(anchor="w", padx=12, pady=(6,2))
        self.conv_result_lbl = tk.Label(prev, text="",
                                         font=("Segoe UI",10), fg=TEXT,
                                         bg=SURFACE, wraplength=800, justify="left")
        self.conv_result_lbl.pack(anchor="w", padx=14, pady=(0,10))

    def _clear_log(self):
        self.log1.configure(state="normal")
        self.log1.delete("1.0","end")
        self.log1.configure(state="disabled")

    # ── EXTRACT TAB LOGIC ─────────────────────────────────────
    def _browse(self):
        f = filedialog.askopenfilename(
            title="Dooro Faylka",
            filetypes=[("All Files","*.*"),
                       ("Archives","*.zip *.rar *.7z *.tar *.gz *.bz2 *.xz *.tgz"),
                       ("Documents","*.pdf *.docx *.xlsx *.txt"),
                       ("Images","*.png *.jpg *.jpeg *.gif *.bmp *.webp"),
                       ("Videos","*.mp4 *.mkv *.avi *.mov")])
        if f: self._load_file(f)

    def _load_file(self, path):
        path = path.strip('"').strip("'")
        if not os.path.exists(path): self._log(f"Faylka ma jiro: {path}","err"); return
        self.current_file = path
        self._cancel = False
        name  = os.path.basename(path)
        size  = os.path.getsize(path)
        ftype = get_type(path)
        mod   = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
        icon  = TYPE_ICONS.get(ftype,"📁")
        color = TYPE_COLORS.get(ftype,TEXT2)
        self.type_lbl.configure(text=icon, fg=color)
        self.name_lbl.configure(text=name)
        self.meta_lbl.configure(text=f"Nooca: {ftype.upper()}  |  Cabbirka: {fmt_size(size)}  |  {mod}")
        self.drop_lbl.configure(text=f"✅  {name}", fg=GREEN)
        self.extract_dest = str(Path(path).parent)
        self.dest_var.set(self.extract_dest)
        self._set_prog(0,"")
        self.contents_list.delete(0,"end")
        if is_archive(ftype):
            self.contents_label.configure(text="📂 Waxa ku jira:")
            threading.Thread(target=self._preview_archive,
                             args=(path,ftype), daemon=True).start()
        else:
            self.contents_label.configure(text="")
        self._log(f"Waa la soo galiyay: {name}  ({fmt_size(size)})","ok")

    def _preview_archive(self, path, ftype):
        entries = []
        try:
            if ftype == "zip":
                with zipfile.ZipFile(path) as z:
                    for i in z.infolist(): entries.append((i.filename,i.file_size))
            elif ftype == "7z" and HAS_7Z:
                with py7zr.SevenZipFile(path,mode="r") as z:
                    for n,i in z.list(): entries.append((n,getattr(i,'uncompressed',0) or 0))
            elif ftype in {"tar","tar.gz","tar.bz2","tar.xz"}:
                with tarfile.open(path,"r:*") as t:
                    for m in t.getmembers(): entries.append((m.name,m.size))
            elif ftype == "rar" and HAS_RAR:
                with rarfile.RarFile(path) as r:
                    for i in r.infolist(): entries.append((i.filename,i.file_size))
        except Exception as ex:
            self._log(f"Preview khalad: {ex}","warn"); return
        self.after(0, lambda: self.contents_label.configure(
            text=f"📂 Waxa ku jira — {len(entries)} file(s):"))
        for name,size in entries[:200]:
            line = f"  {fmt_size(size):>10}   {name}"
            self.after(0, lambda l=line: self.contents_list.insert("end",l))

    def _choose_dest(self):
        d = filedialog.askdirectory(title="Dooro meesha la furi doono")
        if d: self.extract_dest=d; self.dest_var.set(d)

    def _reset_dest(self):
        if self.current_file:
            self.extract_dest = str(Path(self.current_file).parent)
            self.dest_var.set(self.extract_dest)

    def _start_extract(self):
        if not self.current_file:
            messagebox.showwarning("Fayl Ma Jiro","Marka hore fayl soo gali!"); return
        ftype = get_type(self.current_file)
        if not is_archive(ftype):
            self._log("Noocani ma aha archive. Isticmaal Convert baddalkeeda.","warn"); return
        dest = self.dest_var.get().strip()
        if not dest or dest.startswith("("): dest = str(Path(self.current_file).parent)
        os.makedirs(dest, exist_ok=True)
        self.extract_dest = dest; self._cancel = False
        self._set_prog(0,"")
        threading.Thread(target=self._do_extract,
                         args=(self.current_file,ftype,dest), daemon=True).start()

    def _do_extract(self, path, ftype, dest):
        name = os.path.basename(path)
        self._log(f"Bilaabaya furista: {name} → {dest}","info")
        try:
            if ftype=="zip":      self._extr_zip(path,dest)
            elif ftype=="7z" and HAS_7Z: self._extr_7z(path,dest)
            elif ftype in {"tar","tar.gz","tar.bz2","tar.xz"}: self._extr_tar(path,dest)
            elif ftype in {"gz","bz2","xz"}: self._extr_single(path,dest,ftype)
            elif ftype=="rar" and HAS_RAR: self._extr_rar(path,dest)
            else: shutil.unpack_archive(path,dest); self._set_prog(100,"")
            if not self._cancel:
                self.after(0, lambda: self._set_prog(100,"✅ Dhammaystay!"))
                self._log(f"✅ Furista waa dhammaatay! → {dest}","ok")
                self.after(0, lambda: messagebox.showinfo("✅ Guul!",
                    f"Faylku waa si fiican u furmay!\nMeesha: {dest}"))
        except Exception as ex:
            self._log(f"❌ Khalad: {ex}","err")
            self.after(0, lambda e=ex: messagebox.showerror("Khalad!",str(e)))

    def _extr_zip(self, path, dest):
        with zipfile.ZipFile(path,"r") as z:
            members = z.infolist(); total=len(members)
            for i,m in enumerate(members):
                if self._cancel: break
                z.extract(m,dest)
                self.after(0,lambda p=(i+1)/total*100,n=m.filename:
                    self._set_prog(p,f"Furaaya: {n[:70]}"))

    def _extr_7z(self, path, dest):
        self._set_prog(10,"Furaaya 7Z … fadlan sug")
        with py7zr.SevenZipFile(path,mode="r") as z: z.extractall(path=dest)
        self._set_prog(100,"")

    def _extr_tar(self, path, dest):
        with tarfile.open(path,"r:*") as t:
            members=t.getmembers(); total=len(members)
            for i,m in enumerate(members):
                if self._cancel: break
                t.extract(m,dest,filter="data")
                self.after(0,lambda p=(i+1)/total*100,n=m.name:
                    self._set_prog(p,f"Furaaya: {n[:70]}"))

    def _extr_single(self, path, dest, ftype):
        import gzip,bz2,lzma
        openers={"gz":gzip.open,"bz2":bz2.open,"xz":lzma.open}
        out=os.path.join(dest,Path(path).stem); size=os.path.getsize(path); read=0
        with openers[ftype](path,"rb") as fi, open(out,"wb") as fo:
            while True:
                if self._cancel: break
                c=fi.read(1024*1024)
                if not c: break
                fo.write(c); read+=len(c)
                self.after(0,lambda p=min(99,read/size*100) if size else 50:
                    self._set_prog(p,f"Furaaya: {fmt_size(read)}"))

    def _extr_rar(self, path, dest):
        with rarfile.RarFile(path) as r:
            members=r.infolist(); total=len(members)
            for i,m in enumerate(members):
                if self._cancel: break
                r.extract(m,dest)
                self.after(0,lambda p=(i+1)/total*100,n=m.filename:
                    self._set_prog(p,f"Furaaya: {n[:70]}"))

    def _on_drop(self, event):
        """Handle real drag & drop events from tkinterdnd2"""
        raw = event.data.strip()
        # tkinterdnd2 may wrap paths in braces if they contain spaces
        paths = []
        if raw.startswith("{"):
            import re
            paths = re.findall(r'\{([^}]+)\}|([^\s]+)', raw)
            paths = [p[0] or p[1] for p in paths]
        else:
            paths = raw.split()
        if paths:
            self._load_file(paths[0])

    def _cancel_extract(self):
        self._cancel=True; self._set_prog(0,"Waa la joojiyay.")
        self._log("⛔ Waa la joojiyay!","warn")

    def _open_file(self):
        if not self.current_file:
            messagebox.showwarning("Fayl Ma Jiro","Marka hore fayl soo gali!"); return
        try: os.startfile(self.current_file); self._log(f"Waa la furay: {os.path.basename(self.current_file)}","ok")
        except Exception as ex: self._log(f"Furin kuma guulaysan: {ex}","err")

    def _open_dest(self):
        d = self.dest_var.get().strip()
        if d and os.path.isdir(d): os.startfile(d)
        elif self.current_file: os.startfile(str(Path(self.current_file).parent))

    def _clear(self):
        self.current_file=None
        self.type_lbl.configure(text=""); self.name_lbl.configure(text="Fayl ma jiro weli")
        self.meta_lbl.configure(text=""); self.contents_label.configure(text="")
        self.contents_list.delete(0,"end")
        self.drop_lbl.configure(
            text="🗂️   Drag & Drop Archive Here   —   or   Click to Browse",fg=ACCENT_LT)
        self.dest_var.set("(Isla folder-ka faylka)"); self._set_prog(0,"")

    # ── CONVERT TAB LOGIC ─────────────────────────────────────
    def _conv_browse(self):
        f = filedialog.askopenfilename(
            title="Dooro Faylka Aad Badali Rabto",
            filetypes=[("All Files","*.*"),
                       ("Images","*.png *.jpg *.jpeg *.bmp *.webp *.gif *.tiff *.ico"),
                       ("Documents","*.docx *.txt *.md *.pdf"),
                       ("Spreadsheets","*.xlsx *.xls *.csv"),
                       ("Archives","*.zip *.7z *.tar *.gz")])
        if not f: return
        self.conv_src_var.set(f)
        ftype = get_type(f)
        size  = fmt_size(os.path.getsize(f))
        icon  = TYPE_ICONS.get(ftype,"📁")
        self.conv_info_lbl.configure(
            text=f"{icon}  {os.path.basename(f)}  |  {ftype.upper()}  |  {size}",
            fg=TYPE_COLORS.get(ftype,TEXT2))
        # Populate dropdown
        options = CONVERSIONS.get(ftype, [])
        if options:
            labels = [o[0] for o in options]
            self.conv_dropdown["values"] = labels
            self.conv_type_var.set(labels[0])
            self._log(f"Faylka badalka: {os.path.basename(f)}  ({ftype.upper()})","info")
        else:
            self.conv_dropdown["values"] = ["— noocaan badalin karo —"]
            self.conv_type_var.set("— noocaan badalin karo —")
            self._log(f"Nooca '{ftype}' waa la badali karaa oo kaliya formats kooban","warn")
        # Auto set dest
        self.conv_dest_var.set(str(Path(f).parent))

    def _conv_choose_dest(self):
        d = filedialog.askdirectory(title="Dooro meesha la kaydinayo")
        if d: self.conv_dest_var.set(d)

    def _conv_open_dest(self):
        d = self.conv_dest_var.get().strip()
        if d and os.path.isdir(d): os.startfile(d)

    def _start_convert(self):
        src = self.conv_src_var.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning("Fayl Ma Jiro","Marka hore fayl soo dooro!"); return
        label = self.conv_type_var.get()
        if "—" in label:
            messagebox.showwarning("Nooc","Fadlan nooca badalka dooro!"); return
        ftype  = get_type(src)
        opts   = CONVERSIONS.get(ftype, [])
        ext    = None
        for lbl, e in opts:
            if lbl == label: ext = e; break
        if not ext:
            messagebox.showwarning("Khalad","Noocaan la aqoonsan waayo!"); return
        dest_dir = self.conv_dest_var.get().strip()
        if not dest_dir or dest_dir.startswith("("): dest_dir = str(Path(src).parent)
        os.makedirs(dest_dir, exist_ok=True)
        out_name = Path(src).stem + ("" if not ext.startswith(".") else ext)
        # Handle tar.gz
        if ext == ".tar.gz": out_name = Path(src).stem + ".tar.gz"
        out_path = os.path.join(dest_dir, out_name)
        self._cancel = False
        self._set_prog(0,"")
        self.conv_result_lbl.configure(text="Badalayaa …", fg=YELLOW)
        threading.Thread(target=self._do_convert,
                         args=(src, out_path, ftype, ext), daemon=True).start()

    def _do_convert(self, src, out_path, ftype, ext):
        name = os.path.basename(src)
        self._log(f"Bilaabaya badalka: {name} → {os.path.basename(out_path)}","info")
        try:
            # ── IMAGE CONVERSIONS ─────────────────────────────
            if ftype == "image":
                self._convert_image(src, out_path, ext)

            # ── DOCX → PDF ────────────────────────────────────
            elif ftype == "word" and ext == ".pdf":
                self._convert_docx_pdf(src, out_path)

            # ── DOCX → TXT ────────────────────────────────────
            elif ftype == "word" and ext == ".txt":
                self._convert_docx_txt(src, out_path)

            # ── CSV → XLSX ────────────────────────────────────
            elif ftype == "csv" and ext == ".xlsx":
                self._convert_csv_xlsx(src, out_path)

            # ── CSV → JSON / TXT ──────────────────────────────
            elif ftype == "csv" and ext in (".json",".txt"):
                self._convert_csv_other(src, out_path, ext)

            # ── XLSX → CSV / JSON / TXT ───────────────────────
            elif ftype == "excel":
                self._convert_xlsx_other(src, out_path, ext)

            # ── PDF → TXT ─────────────────────────────────────
            elif ftype == "pdf" and ext == ".txt":
                self._convert_pdf_txt(src, out_path)

            # ── ARCHIVE CONVERSIONS ───────────────────────────
            elif ftype == "zip" and ext == ".7z":
                self._convert_zip_7z(src, out_path)
            elif ftype == "zip" and ext == ".tar.gz":
                self._convert_zip_targz(src, out_path)
            elif ftype == "7z" and ext == ".zip":
                self._convert_7z_zip(src, out_path)
                
            # ── CODE & TEXT & OTHER ───────────────────────────
            elif ftype in ("text", "html", "python", "javascript", "json", "xml", "css", "code", "sql", "other"):
                if ext == ".pdf":
                    if ftype == "html":
                        try:
                            from xhtml2pdf import pisa
                            self._set_prog(20, "PDF HTML Rendering…")
                            with open(src, "r", encoding="utf-8", errors="replace") as f_in:
                                source_html = f_in.read()
                            with open(out_path, "w+b") as result_file:
                                pisa_status = pisa.CreatePDF(source_html, dest=result_file)
                            self._set_prog(100, "")
                            if pisa_status.err:
                                raise Exception("Khalad dhinaca xhtml2pdf rendering")
                        except Exception as ex:
                            self._log(f"Fallback to text PDF: {ex}", "warn")
                            # Fallback if xhtml2pdf fails
                            try:
                                from bs4 import BeautifulSoup
                                with open(src, "r", encoding="utf-8", errors="replace") as f_in:
                                    content = BeautifulSoup(f_in.read(), "html.parser").get_text(separator="\n", strip=True)
                                tmp = out_path + ".tmp.txt"
                                with open(tmp, "w", encoding="utf-8") as f_out:
                                    f_out.write(content)
                                self._convert_txt_pdf(tmp, out_path)
                                os.remove(tmp)
                            except Exception:
                                self._convert_txt_pdf(src, out_path)
                    else:
                        self._convert_txt_pdf(src, out_path)
                elif ext == ".html":
                    self._convert_txt_html(src, out_path)
                elif ext in (".txt", ".md", ".json", ".csv"):
                    self._set_prog(50, f"Baddalayaa oo ka dhigaya {ext.upper()} …")
                    with open(src, "r", encoding="utf-8", errors="replace") as f_in:
                        content = f_in.read()
                    
                    if ftype == "html" and ext in (".txt", ".md"):
                        try:
                            from bs4 import BeautifulSoup
                            content = BeautifulSoup(content, "html.parser").get_text(separator="\n", strip=True)
                        except Exception:
                            pass
                            
                    with open(out_path, "w", encoding="utf-8") as f_out:
                        f_out.write(content)
                    self._set_prog(100, "")
                else:
                    raise NotImplementedError(f"Badalkan '{ext}' wali la sameyn waayay")
            else:
                raise NotImplementedError(f"Badalkan '{ext}' wali la sameyn waayay")

            if not self._cancel:
                self._set_prog(100,"✅ Badalka waa dhammaatay!")
                size_out = fmt_size(os.path.getsize(out_path)) if os.path.exists(out_path) else "?"
                self._log(f"✅ Badalka waa guulaystay! → {out_path}  ({size_out})","ok")
                self.after(0, lambda: self.conv_result_lbl.configure(
                    text=f"✅  Waa la baddalay!\n📄  {out_path}\n📦  Cabbirka: {size_out}",
                    fg=GREEN))
                self.after(0, lambda: messagebox.showinfo("✅ Guul!",
                    f"Badalka waa dhammaatay!\n\nFaylka cusub:\n{out_path}"))
        except Exception as ex:
            self._log(f"❌ Khalad badalka: {ex}","err")
            self.after(0, lambda e=ex: self.conv_result_lbl.configure(
                text=f"❌ Khalad: {e}", fg=RED))
            self.after(0, lambda e=ex: messagebox.showerror("Khalad!", str(e)))

    # ── CONVERTERS ────────────────────────────────────────────
    def _convert_image(self, src, out, ext):
        if not HAS_PIL: raise RuntimeError("Pillow ma diyaar — pip install pillow")
        self._set_prog(20,"Sawirka waa la xidxidayaa …")
        img = Image.open(src)

        # Resize if needed
        resize_pct = self.resize_var.get()
        if resize_pct != 100:
            w = int(img.width * resize_pct / 100)
            h = int(img.height * resize_pct / 100)
            img = img.resize((w,h), Image.LANCZOS)

        quality = self.quality_var.get()
        self._set_prog(60,"Badalaaya …")

        if ext == ".ico":
            img.save(out, format="ICO", sizes=[(256,256),(128,128),(64,64),(32,32),(16,16)])
        elif ext == ".pdf":
            # Image to PDF
            paper = self.pdf_size_var.get()
            paper_sizes = {"A4":(595,842),"A3":(842,1191),
                           "Letter":(612,792),"Legal":(612,1008)}
            pw,ph = paper_sizes.get(paper,(595,842))
            if img.mode in ("RGBA","P","LA"): img = img.convert("RGB")
            # Scale image to fit page
            ratio = min(pw/img.width, ph/img.height)
            nw,nh = int(img.width*ratio), int(img.height*ratio)
            img = img.resize((nw,nh), Image.LANCZOS)
            bg = Image.new("RGB",(pw,ph),(255,255,255))
            bg.paste(img,((pw-nw)//2,(ph-nh)//2))
            bg.save(out,"PDF")
        elif ext in (".jpg",".jpeg"):
            if img.mode in ("RGBA","P","LA"): img = img.convert("RGB")
            img.save(out, format="JPEG", quality=quality, optimize=True)
        elif ext == ".webp":
            img.save(out, format="WEBP", quality=quality)
        elif ext == ".tiff":
            img.save(out, format="TIFF")
        else:
            img.save(out)
        self._set_prog(100,"")

    def _convert_docx_pdf(self, src, out):
        # Try docx2pdf (needs MS Word on Windows)
        try:
            from docx2pdf import convert
            self._set_prog(30,"DOCX → PDF badalayaa …")
            convert(src, out)
            self._set_prog(100,"")
        except Exception:
            # Fallback: extract text and make PDF with fpdf2
            self._convert_docx_txt_then_pdf(src, out)

    def _convert_docx_txt_then_pdf(self, src, out):
        if not HAS_DOCX: raise RuntimeError("python-docx ma diyaar — pip install python-docx")
        doc = Document(src)
        lines = [p.text for p in doc.paragraphs]
        self._set_prog(40,"Qoraalka la saaraya …")
        txt = "\n".join(lines)
        tmp = out + ".tmp.txt"
        with open(tmp,"w",encoding="utf-8") as f: f.write(txt)
        self._convert_txt_pdf(tmp, out)
        os.remove(tmp)

    def _convert_docx_txt(self, src, out):
        if not HAS_DOCX: raise RuntimeError("python-docx ma diyaar — pip install python-docx")
        self._set_prog(30,"Qoraalka la saaraya …")
        doc = Document(src)
        lines = [p.text for p in doc.paragraphs]
        with open(out,"w",encoding="utf-8") as f: f.write("\n".join(lines))
        self._set_prog(100,"")

    def _convert_txt_pdf(self, src, out):
        if HAS_FPDF:
            import textwrap
            self._set_prog(20,"TXT → PDF badalayaa …")
            pdf = FPDF()
            pdf.set_auto_page_break(auto=True, margin=15)
            paper = self.pdf_size_var.get()
            fmt_map = {"A4":"A4","A3":"A3","Letter":(216,279),"Legal":(216,356)}
            fmt = fmt_map.get(paper,"A4")
            pdf.add_page(format=fmt if isinstance(fmt,tuple) else fmt)
            pdf.set_font("Helvetica", size=11)
            with open(src,"r",encoding="utf-8",errors="replace") as f:
                lines = f.readlines()
            total = len(lines)
            for i,line in enumerate(lines):
                if self._cancel: break
                text = line.rstrip()
                # Wrap long words so fpdf doesn't crash on base64 strings
                wrapped_lines = textwrap.wrap(text, width=95, break_long_words=True, replace_whitespace=False) or [""]
                for w_line in wrapped_lines:
                    pdf.multi_cell(0, 6, w_line)
                if i % 50 == 0:
                    self.after(0,lambda p=20+i/total*75:self._set_prog(p,"Qoraaya PDF…"))
            pdf.output(out)
            self._set_prog(100,"")
        elif HAS_RL:
            self._convert_txt_pdf_reportlab(src, out)
        else:
            raise RuntimeError("fpdf2 ama reportlab ma diyaar — pip install fpdf2")

    def _convert_txt_pdf_reportlab(self, src, out):
        from reportlab.lib.pagesizes import A4, letter
        from reportlab.platypus import SimpleDocTemplate, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        self._set_prog(20,"TXT → PDF (reportlab) …")
        doc = SimpleDocTemplate(out, pagesize=A4)
        styles = getSampleStyleSheet()
        story = []
        with open(src,"r",encoding="utf-8",errors="replace") as f:
            for line in f:
                story.append(Paragraph(line.strip() or "&nbsp;", styles["Normal"]))
        doc.build(story)
        self._set_prog(100,"")

    def _convert_txt_html(self, src, out):
        self._set_prog(30,"TXT → HTML badalayaa …")
        with open(src,"r",encoding="utf-8",errors="replace") as f: txt = f.read()
        html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>{Path(src).name}</title>
<style>body{{font-family:monospace;padding:30px;background:#1e1e1e;color:#d4d4d4;white-space:pre-wrap;}}</style>
</head><body>{txt.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')}</body></html>"""
        with open(out,"w",encoding="utf-8") as f: f.write(html)
        self._set_prog(100,"")

    def _convert_csv_xlsx(self, src, out):
        if not (HAS_PD and HAS_XL):
            raise RuntimeError("pandas ama openpyxl ma diyaar — pip install pandas openpyxl")
        self._set_prog(30,"CSV → XLSX badalayaa …")
        df = pd.read_csv(src, encoding="utf-8-sig", low_memory=False)
        self._set_prog(70,"Keydinaya …")
        df.to_excel(out, index=False, engine="openpyxl")
        self._set_prog(100,"")

    def _convert_csv_other(self, src, out, ext):
        if not HAS_PD: raise RuntimeError("pandas ma diyaar — pip install pandas")
        self._set_prog(30,"CSV waa la xidxidayaa …")
        df = pd.read_csv(src, encoding="utf-8-sig")
        self._set_prog(60,"Badalayaa …")
        if ext == ".json": df.to_json(out, orient="records", indent=2)
        elif ext == ".txt": df.to_csv(out, sep="\t", index=False)
        self._set_prog(100,"")

    def _convert_xlsx_other(self, src, out, ext):
        if not HAS_PD: raise RuntimeError("pandas ma diyaar — pip install pandas")
        self._set_prog(30,"XLSX waa la xidxidayaa …")
        df = pd.read_excel(src, engine="openpyxl")
        self._set_prog(65,"Badalayaa …")
        if ext == ".csv": df.to_csv(out, index=False)
        elif ext == ".json": df.to_json(out, orient="records", indent=2)
        elif ext == ".txt": df.to_csv(out, sep="\t", index=False)
        self._set_prog(100,"")

    def _convert_pdf_txt(self, src, out):
        try:
            import pypdf
            self._set_prog(20,"PDF → TXT waa la soo sarayaa …")
            reader = pypdf.PdfReader(src)
            pages = reader.pages; total = len(pages); text = []
            for i,page in enumerate(pages):
                if self._cancel: break
                text.append(page.extract_text() or "")
                self.after(0,lambda p=(i+1)/total*90:self._set_prog(p,f"Page {i+1}/{total}"))
            with open(out,"w",encoding="utf-8") as f: f.write("\n\n".join(text))
            self._set_prog(100,"")
        except ImportError:
            raise RuntimeError("pypdf ma diyaar — pip install pypdf")

    def _convert_zip_7z(self, src, out):
        if not HAS_7Z: raise RuntimeError("py7zr ma diyaar — pip install py7zr")
        self._set_prog(15,"ZIP → 7Z: furaaya ZIP …")
        tmp = out + "_tmp_extract"
        with zipfile.ZipFile(src,"r") as z: z.extractall(tmp)
        self._set_prog(50,"Abuuraya 7Z …")
        with py7zr.SevenZipFile(out,mode="w") as z: z.writeall(tmp, ".")
        shutil.rmtree(tmp, ignore_errors=True)
        self._set_prog(100,"")

    def _convert_zip_targz(self, src, out):
        self._set_prog(15,"ZIP → TAR.GZ: furaaya ZIP …")
        tmp = out + "_tmp_extract"
        with zipfile.ZipFile(src,"r") as z: z.extractall(tmp)
        self._set_prog(50,"Abuuraya TAR.GZ …")
        with tarfile.open(out,"w:gz") as t: t.add(tmp, arcname=".")
        shutil.rmtree(tmp, ignore_errors=True)
        self._set_prog(100,"")

    def _convert_7z_zip(self, src, out):
        if not HAS_7Z: raise RuntimeError("py7zr ma diyaar — pip install py7zr")
        self._set_prog(15,"7Z → ZIP: furaaya 7Z …")
        tmp = out + "_tmp_extract"
        with py7zr.SevenZipFile(src,mode="r") as z: z.extractall(tmp)
        self._set_prog(55,"Abuuraya ZIP …")
        with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
            for root,dirs,files in os.walk(tmp):
                for file in files:
                    fp = os.path.join(root,file)
                    z.write(fp, os.path.relpath(fp,tmp))
        shutil.rmtree(tmp, ignore_errors=True)
        self._set_prog(100,"")

    # ══════════════════════════════════════════════════════════
    # TAB 3 — BATCH EXTRACT
    # ══════════════════════════════════════════════════════════
    def _build_tab_batch(self, parent):
        tk.Label(parent, text="📂 Xog Badan — Extract Archives Badan Hal Mar",
                 font=("Segoe UI",13,"bold"), fg=ACCENT_LT, bg=BG).pack(anchor="w", padx=14, pady=(12,4))
        tk.Label(parent, text="Dooro archives badan oo dhammaantood hal mar furi",
                 font=("Segoe UI",9), fg=TEXT2, bg=BG).pack(anchor="w", padx=16, pady=(0,8))

        # File list
        lf = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        lf.pack(fill="both", expand=True, padx=12, pady=4)

        hdr = tk.Frame(lf, bg=SURFACE)
        hdr.pack(fill="x", padx=8, pady=(6,2))
        tk.Label(hdr, text="Liiska Faylasha:", font=("Segoe UI",9,"bold"), fg=TEXT2, bg=SURFACE).pack(side="left")
        self._btn(hdr, "➕ Kudar", self._batch_add, ACCENT, "white", padx=10, pady=3).pack(side="right", padx=2)
        self._btn(hdr, "🗑️ Ka Saar", self._batch_remove, SURFACE2, RED, padx=10, pady=3).pack(side="right", padx=2)
        self._btn(hdr, "❌ Nadiifi", self._batch_clear, SURFACE2, TEXT2, padx=10, pady=3).pack(side="right", padx=2)

        lb_frame = tk.Frame(lf, bg=SURFACE2)
        lb_frame.pack(fill="both", expand=True, padx=8, pady=(2,8))
        bsb = tk.Scrollbar(lb_frame); bsb.pack(side="right", fill="y")
        self.batch_list = tk.Listbox(lb_frame, bg=SURFACE2, fg=TEXT,
                                     selectbackground=ACCENT, font=("Consolas",9),
                                     relief="flat", bd=0, selectmode="extended",
                                     yscrollcommand=bsb.set)
        self.batch_list.pack(fill="both", expand=True)
        bsb.config(command=self.batch_list.yview)

        if HAS_DND:
            self.batch_list.drop_target_register(DND_FILES)
            self.batch_list.dnd_bind("<<Drop>>", self._batch_on_drop)

        # Destination
        df = tk.Frame(parent, bg=BG, padx=12); df.pack(fill="x", pady=4)
        tk.Label(df, text="📂 Meesha la furi doono (Output Folder):",
                 font=("Segoe UI",10,"bold"), fg=TEXT2, bg=BG).pack(anchor="w")
        dr = tk.Frame(df, bg=BG); dr.pack(fill="x", pady=3)
        self.batch_dest_var = tk.StringVar(value="(Isla folder-ka fayl kasta)")
        tk.Entry(dr, textvariable=self.batch_dest_var, font=("Segoe UI",10),
                 fg=TEXT, bg=SURFACE2, relief="flat", bd=6,
                 insertbackground=TEXT).pack(side="left", fill="x", expand=True, ipady=5)
        self._btn(dr, "📁 Dooro", self._batch_choose_dest, SURFACE2, TEXT, padx=12, pady=6).pack(side="left", padx=4)

        # Buttons
        bf = tk.Frame(parent, bg=BG, padx=12, pady=8); bf.pack(fill="x")
        self._btn(bf, "🚀 FURI DHAMMAAN / EXTRACT ALL", self._batch_extract_all,
                  ACCENT, "white", font=("Segoe UI",12,"bold"), padx=22, pady=10).pack(side="left", padx=(0,8))
        self._btn(bf, "⛔ Jooji", self._cancel_extract, SURFACE2, RED,
                  padx=14, pady=10).pack(side="left")

        # Stats label
        self.batch_stats_lbl = tk.Label(parent, text="",
                                         font=("Segoe UI",9), fg=TEXT2, bg=BG)
        self.batch_stats_lbl.pack(anchor="w", padx=14)

    def _batch_add(self):
        files = filedialog.askopenfilenames(
            title="Dooro Archives",
            filetypes=[("Archives", "*.zip *.rar *.7z *.tar *.gz *.bz2 *.xz *.tgz"),
                       ("All Files", "*.*")])
        for f in files:
            if f not in self.batch_list.get(0, "end"):
                self.batch_list.insert("end", f)
        self._update_batch_stats()

    def _batch_remove(self):
        for i in reversed(self.batch_list.curselection()):
            self.batch_list.delete(i)
        self._update_batch_stats()

    def _batch_clear(self):
        self.batch_list.delete(0, "end")
        self._update_batch_stats()

    def _batch_on_drop(self, event):
        raw = event.data.strip()
        import re
        if raw.startswith("{"):
            paths = re.findall(r'\{([^}]+)\}|([^\s]+)', raw)
            paths = [p[0] or p[1] for p in paths]
        else:
            paths = raw.split()
        for p in paths:
            p = p.strip('"').strip("'")
            if os.path.isfile(p) and is_archive(get_type(p)):
                if p not in self.batch_list.get(0, "end"):
                    self.batch_list.insert("end", p)
        self._update_batch_stats()

    def _update_batch_stats(self):
        items = self.batch_list.get(0, "end")
        total_size = sum(os.path.getsize(f) for f in items if os.path.exists(f))
        self.batch_stats_lbl.configure(
            text=f"  📦 {len(items)} archive(s)   |   Wadarta: {fmt_size(total_size)}")

    def _batch_choose_dest(self):
        d = filedialog.askdirectory(title="Dooro meesha la furi doono")
        if d: self.batch_dest_var.set(d)

    def _batch_extract_all(self):
        items = list(self.batch_list.get(0, "end"))
        if not items:
            messagebox.showwarning("Liis Madhan", "Marka hore archives ku dar liiska!"); return
        self._cancel = False
        threading.Thread(target=self._do_batch_extract, args=(items,), daemon=True).start()

    def _do_batch_extract(self, files):
        total = len(files)
        done = 0; errors = 0
        for idx, path in enumerate(files):
            if self._cancel: break
            if not os.path.exists(path):
                self._log(f"Ma jiro: {os.path.basename(path)}", "warn"); errors += 1; continue
            ftype = get_type(path)
            if not is_archive(ftype):
                self._log(f"Nooc lama garanayn: {os.path.basename(path)}", "warn"); errors += 1; continue

            dest_root = self.batch_dest_var.get().strip()
            if not dest_root or dest_root.startswith("("):
                dest = str(Path(path).parent / Path(path).stem)
            else:
                dest = os.path.join(dest_root, Path(path).stem)

            os.makedirs(dest, exist_ok=True)
            self.after(0, lambda i=idx, t=total, n=os.path.basename(path):
                self._set_prog(i/t*100, f"[{i+1}/{t}] Furaaya: {n[:60]}"))
            self._log(f"[{idx+1}/{total}] Furaaya: {os.path.basename(path)}", "info")

            try:
                if ftype == "zip":   self._extr_zip(path, dest)
                elif ftype == "7z" and HAS_7Z: self._extr_7z(path, dest)
                elif ftype in {"tar","tar.gz","tar.bz2","tar.xz"}: self._extr_tar(path, dest)
                elif ftype in {"gz","bz2","xz"}: self._extr_single(path, dest, ftype)
                elif ftype == "rar" and HAS_RAR: self._extr_rar(path, dest)
                else: shutil.unpack_archive(path, dest)
                self._log(f"  ✅ Guul → {dest}", "ok")
                done += 1
            except Exception as ex:
                self._log(f"  ❌ Khalad: {ex}", "err"); errors += 1

        self.after(0, lambda: self._set_prog(100, f"✅ Dhammaystay! {done}/{total} guul, {errors} khalad"))
        self._log(f"🎉 Batch dhammaatay: {done} guul, {errors} khalad, {total-done-errors} la joojiyay", "ok")
        self.after(0, lambda: messagebox.showinfo("Batch Dhammaatay",
            f"Natiijada:\n✅ Guul: {done}\n❌ Khalad: {errors}\n📦 Wadarta: {total}"))

    # ══════════════════════════════════════════════════════════
    # TAB 4 — SETTINGS (Context Menu + Info)
    # ══════════════════════════════════════════════════════════
    def _build_tab_settings(self, parent):
        tk.Label(parent, text="⚙️ Dejinta — Settings",
                 font=("Segoe UI",14,"bold"), fg=ACCENT_LT, bg=BG).pack(anchor="w", padx=14, pady=(14,4))

        # ── Windows Context Menu ─────────────────────────────
        cm = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        cm.pack(fill="x", padx=12, pady=8)
        tk.Label(cm, text="🖱️ Windows Right-Click Menu (Context Menu)",
                 font=("Segoe UI",11,"bold"), fg=TEXT, bg=SURFACE).pack(anchor="w", padx=12, pady=(10,4))
        tk.Label(cm, text="Faylasha archive-ka ku dhufo midigta, FileOpener-ka ku dar menu-ga",
                 font=("Segoe UI",9), fg=TEXT2, bg=SURFACE).pack(anchor="w", padx=14)

        self.ctxmenu_status = tk.Label(cm, text="",
                                        font=("Segoe UI",9), fg=TEXT2, bg=SURFACE)
        self.ctxmenu_status.pack(anchor="w", padx=14, pady=(2,6))
        self._check_context_menu_status()

        btnf = tk.Frame(cm, bg=SURFACE); btnf.pack(anchor="w", padx=12, pady=(0,12))
        self._btn(btnf, "✅ Ku Dar Context Menu", self._install_context_menu,
                  GREEN, "white", padx=14, pady=7).pack(side="left", padx=(0,8))
        self._btn(btnf, "❌ Ka Saar Context Menu", self._uninstall_context_menu,
                  SURFACE2, RED, padx=14, pady=7).pack(side="left")

        # ── Library Status ───────────────────────────────────
        ls = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        ls.pack(fill="x", padx=12, pady=4)
        tk.Label(ls, text="📦 Libraries Status",
                 font=("Segoe UI",11,"bold"), fg=TEXT, bg=SURFACE).pack(anchor="w", padx=12, pady=(10,4))

        libs = [
            ("tkinterdnd2", HAS_DND,  "Drag & Drop support"),
            ("py7zr",       HAS_7Z,   "7Z/7Zip archives"),
            ("rarfile",     HAS_RAR,  "RAR archives (also needs unrar.exe)"),
            ("Pillow",      HAS_PIL,  "Image conversion"),
            ("pandas",      HAS_PD,   "CSV/XLSX data formats"),
            ("openpyxl",    HAS_XL,   "Excel files"),
            ("python-docx", HAS_DOCX, "DOCX documents"),
            ("fpdf2",       HAS_FPDF, "PDF writing"),
        ]
        for lib, ok, desc in libs:
            row = tk.Frame(ls, bg=SURFACE); row.pack(fill="x", padx=14, pady=1)
            ico = "✅" if ok else "❌"
            clr = GREEN if ok else RED
            tk.Label(row, text=f"{ico}  {lib}", font=("Consolas",9,"bold"),
                     fg=clr, bg=SURFACE, width=18, anchor="w").pack(side="left")
            tk.Label(row, text=f"— {desc}", font=("Segoe UI",9),
                     fg=TEXT2, bg=SURFACE).pack(side="left")

        # Install all button
        ibf = tk.Frame(ls, bg=SURFACE); ibf.pack(anchor="w", padx=12, pady=(8,12))
        self._btn(ibf, "⬇️ Install All Libraries", self._install_libs,
                  ACCENT, "white", padx=14, pady=7).pack(side="left")

    def _check_context_menu_status(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,
                                  r"*\shell\FileOpener", 0, winreg.KEY_READ)
            winreg.CloseKey(key)
            self.ctxmenu_status.configure(
                text="✅ Context Menu waa la dhejiyay", fg=GREEN)
        except FileNotFoundError:
            self.ctxmenu_status.configure(
                text="❌ Context Menu wali lama dhejin", fg=RED)
        except Exception as e:
            self.ctxmenu_status.configure(text=f"⚠️ {e}", fg=YELLOW)

    def _install_context_menu(self):
        try:
            script_path = os.path.abspath(__file__)
            python_exe  = sys.executable
            cmd_value   = f'"{python_exe}" "{script_path}" "%1"'

            # Main key: *\shell\FileOpener
            key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"*\shell\FileOpener")
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "📦 Open with FileOpener")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, python_exe)
            winreg.CloseKey(key)

            # Command subkey
            cmd_key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"*\shell\FileOpener\command")
            winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, cmd_value)
            winreg.CloseKey(cmd_key)

            self._check_context_menu_status()
            messagebox.showinfo("✅ Guul!",
                "Context Menu si fiican ayaa loo dhejiyay!\n"
                "Hadda faylasha ku dhufo midigta oo aad FileOpener.")
        except PermissionError:
            messagebox.showerror("Admin Required",
                "Xaqqa Admin ayaa looga baahan yahay.\n"
                "Bilow FileOpener sidatan:\n"
                "  Right-click run.bat → Run as Administrator")
        except Exception as e:
            messagebox.showerror("Khalad!", str(e))

    def _uninstall_context_menu(self):
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, r"*\shell\FileOpener\command")
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, r"*\shell\FileOpener")
            self._check_context_menu_status()
            messagebox.showinfo("✅", "Context Menu waa la saaray.")
        except FileNotFoundError:
            messagebox.showinfo("Macluumaad", "Context Menu horaan uma dhejsanayn.")
        except PermissionError:
            messagebox.showerror("Admin Required",
                "Run as Administrator si aad u saarto.")
        except Exception as e:
            messagebox.showerror("Khalad!", str(e))

    def _install_libs(self):
        """Open terminal and install all required libraries"""
        cmd = f'"{sys.executable}" -m pip install py7zr rarfile Pillow pandas openpyxl python-docx fpdf2 pypdf tkinterdnd2'
        self._log("Bilaabaya install-ka libraries ...", "info")
        try:
            subprocess.Popen(["powershell", "-NoExit", "-Command", cmd],
                             creationflags=subprocess.CREATE_NEW_CONSOLE)
            messagebox.showinfo("Install",
                "Terminal cusub ayaa la furay.\n"
                "Sug install-ka markuu dhammeeyo, kadibna dib u bilow FileOpener.")
        except Exception as e:
            messagebox.showerror("Khalad!", str(e))


if __name__ == "__main__":
    # Handle command-line argument (from context menu)
    if len(sys.argv) > 1:
        _cli_file = sys.argv[1]
    else:
        _cli_file = None

    app = FileOpener()

    # If launched from context menu with a file argument, load it
    if _cli_file and os.path.exists(_cli_file):
        app.after(300, lambda: app._load_file(_cli_file))

    app.mainloop()
