#!/usr/bin/env python3
"""
Wildcard Editor — A standalone text editor optimized for Stable Diffusion wildcards.
Requires Python 3.8+ with tkinter (standard library).
Run: python wildcard_editor.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, font as tkfont
import os, json, re, time, uuid, shutil, hashlib
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS & DEFAULT SETTINGS
# ─────────────────────────────────────────────────────────────────────────────
APP_NAME = "Wildcard Editor"
CONFIG_PATH = Path.home() / ".wildcard_editor" / "config.json"
CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)

COLORS = {
    "bg0":       "#0d0f14",
    "bg1":       "#13161e",
    "bg2":       "#1a1e28",
    "bg3":       "#222636",
    "bg4":       "#2a2f42",
    "border":    "#2a2f42",
    "border2":   "#353c55",
    "accent":    "#7eb8f7",
    "accent2":   "#a78bfa",
    "accent3":   "#34d399",
    "warn":      "#f59e0b",
    "danger":    "#f87171",
    "text0":     "#e8eaf0",
    "text1":     "#a0a8c0",
    "text2":     "#606880",
    "sel_bg":    "#1e3a5f",
    "wc_hl":     "#2a1f42",   # wildcard highlight bg
    "wc_hl_fg":  "#c4b5fd",   # wildcard highlight fg
    "tab_act":   "#0d0f14",
    "tab_inact": "#13161e",
}

PRESET_COLORS = [
    "#7eb8f7","#a78bfa","#34d399","#f59e0b","#f87171",
    "#fb923c","#e879f9","#22d3ee","#4ade80","#facc15",
    "#f472b6","#818cf8","#94a3b8","#ffffff",
]

DEFAULT_SETTINGS = {
    "wrap_str":        "__",
    "wc_dir":          r"E:\stable-diffusion-webui-reForge\extensions\sd-dynamic-prompts\wildcards",
    "spell_check":     True,
    "font_size":       13,
    "sort_mode":       "name",
    "last_open_dir":   "",
    "word_wrap":       True,
    "window_geometry": "1400x860",
    "sidebar_width":   280,
    "panel_ratio":     0.65,
    "autosave":        False,
}

COMMON_WORDS = set("""
the be to of and a in that have it for not on with he as you do at
this but his by from they we say her she or an will my one all would
there their what so up out if about who get which go me when make can
like time no just him know take people into year your good some could
them see other than then now look only come its over think also back
after use two how our work first well way even new want because any
these give day most us
""".split())


def _contrast_color(hex_color):
    """Return black or white depending on which is more legible against hex_color."""
    try:
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255
        return "#0d0f14" if luminance > 0.5 else "#e8eaf0"
    except Exception:
        return "#e8eaf0"

def _color_tint(hex_color, alpha, base="#13161e"):
    """Blend hex_color over base at given alpha (0..1). Used for subtle row tints."""
    try:
        h = hex_color.lstrip("#")
        fr, fg, fb = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        b = base.lstrip("#")
        br, bg_, bb = int(b[0:2], 16), int(b[2:4], 16), int(b[4:6], 16)
        r = int(br + (fr - br) * alpha)
        g = int(bg_ + (fg - bg_) * alpha)
        b2 = int(bb + (fb - bb) * alpha)
        return f"#{r:02x}{g:02x}{b2:02x}"
    except Exception:
        return base


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG PERSISTENCE
# ─────────────────────────────────────────────────────────────────────────────
def load_config():
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            cfg = {**DEFAULT_SETTINGS, **data}
            return cfg
    except Exception:
        pass
    return dict(DEFAULT_SETTINGS)

def save_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass

def load_tree_state():
    p = CONFIG_PATH.parent / "tree.json"
    try:
        if p.exists():
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {"folders": [], "unsorted": [], "docs": {}}

def save_tree_state(state):
    p = CONFIG_PATH.parent / "tree.json"
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception:
        pass


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────────────────────────────────────
class WildcardEditor:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.cfg = load_config()
        self.root.geometry(self.cfg.get("window_geometry", "1400x860"))
        self.root.minsize(800, 500)
        self.root.configure(bg=COLORS["bg0"])

        # State
        self.tree_state = load_tree_state()
        self.tabs = []          # list of doc_id
        self.active_tab = None  # doc_id
        self.nav_history = []
        self.nav_index = -1
        self.find_matches = []
        self.find_current = -1
        self.find_mode = "normal"  # normal | extended | regex
        self.ctx_target = None   # ("folder"|"doc", id)
        self._drag_data = None
        self._rename_after = None
        self._last_click_item = None
        self._last_click_time = 0
        self._highlight_job = None
        self._status_job = None
        self.word_wrap = tk.BooleanVar(value=self.cfg.get("word_wrap", True))
        self.spell_enabled = self.cfg.get("spell_check", True)
        # Per-tab undo storage: hidden Text widgets that preserve undo stacks
        # Key: doc_id, Value: tk.Text widget (never packed/displayed)
        self._undo_store      = {}    # doc_id → {"stack":[(content,cursor),...], "pos":int}
        self._pinned_tabs     = set() # doc_ids that are pinned
        self._undo_inhibit    = False # True while modifying editor programmatically
        self._undo_generation = 0     # incremented on every _apply_snapshot / tab load
        self._load_count      = 0     # >0 while programmatic loads are in-flight
        self._snap_timer      = None  # after() id for debounced snapshot on typing
        self._snap_gen        = 0     # generation value when snap timer was scheduled

        # Ensure docs dict and unsorted list exist
        if "docs" not in self.tree_state:    self.tree_state["docs"] = {}
        if "folders" not in self.tree_state: self.tree_state["folders"] = []
        if "unsorted" not in self.tree_state: self.tree_state["unsorted"] = []

        self._build_ui()
        self._bind_keys()
        self._refresh_tree()
        self._update_wc_list()

        # Restore tabs
        saved_tabs = self.cfg.get("open_tabs", [])
        saved_active = self.cfg.get("active_tab", None)
        for tid in saved_tabs:
            if tid in self.tree_state["docs"]:
                self._open_tab(tid, push_nav=False)
        if saved_active and saved_active in self.tree_state["docs"]:
            self._switch_tab(saved_active, push_nav=False)
        elif not self.tabs:
            self._new_file()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Defer post-render tasks: line numbers need the widget to be visible
        # before dlineinfo() works, and spell check needs the editor populated.
        self.root.after(150, self._post_startup)

    # ── UI CONSTRUCTION ──────────────────────────────────────────────────────
    def _post_startup(self):
        """Run after the window is fully drawn — dlineinfo() works here."""
        self._update_line_numbers()
        self._apply_scroll_padding()
        # Set the correct spell button state and run if enabled
        if self.spell_enabled:
            self.sb_spell.config(text="Spell ✓", fg=COLORS["accent3"])
            self._run_spell_check()
        else:
            self.sb_spell.config(text="Spell ✗", fg=COLORS["text2"])

    def _build_ui(self):
        self._apply_ttk_styles()
        self._build_toolbar()
        self._build_main()
        self._build_statusbar()

    def _apply_ttk_styles(self):
        style = ttk.Style()
        style.theme_use("default")
        bg, fg, border = COLORS["bg2"], COLORS["text0"], COLORS["border2"]

        style.configure(".",
            background=bg, foreground=fg,
            relief="flat", borderwidth=0,
            font=("Segoe UI", 10))

        style.configure("Treeview",
            background=COLORS["bg1"], foreground=COLORS["text1"],
            fieldbackground=COLORS["bg1"], rowheight=26,
            borderwidth=0, relief="flat")
        style.configure("Treeview.Heading",
            background=COLORS["bg2"], foreground=COLORS["text2"],
            relief="flat", borderwidth=0)
        style.map("Treeview",
            background=[("selected", COLORS["bg3"])],
            foreground=[("selected", COLORS["accent"])])

        style.configure("TPanedwindow", background=COLORS["bg0"])
        style.configure("TSeparator", background=COLORS["border2"])

        style.configure("TScrollbar",
            background=COLORS["bg3"], troughcolor=COLORS["bg1"],
            borderwidth=0, arrowsize=12, relief="flat")
        style.map("TScrollbar", background=[("active", COLORS["bg4"])])

        style.configure("TCombobox",
            background=COLORS["bg3"], foreground=COLORS["text0"],
            fieldbackground=COLORS["bg3"],
            selectbackground=COLORS["bg4"],
            borderwidth=1, relief="flat",
            arrowcolor=COLORS["text2"])

    def _btn(self, parent, text, cmd, fg=None, icon=None, **kw):
        """Toolbar button factory."""
        f = fg or COLORS["text1"]
        b = tk.Label(parent, text=(icon + " " if icon else "") + text,
                     bg=COLORS["bg1"], fg=f,
                     font=("Segoe UI", 10, "bold"),
                     cursor="hand2", padx=9, pady=4,
                     relief="flat", **kw)
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.config(bg=COLORS["bg3"], fg=COLORS["text0"]))
        b.bind("<Leave>", lambda e: b.config(bg=COLORS["bg1"], fg=f))
        return b

    def _build_toolbar(self):
        self.toolbar = tk.Frame(self.root, bg=COLORS["bg1"], height=44,
                                relief="flat", bd=0)
        self.toolbar.pack(side="top", fill="x")
        self.toolbar.pack_propagate(False)

        def sep():
            tk.Frame(self.toolbar, bg=COLORS["border2"], width=1).pack(
                side="left", fill="y", padx=4, pady=6)

        # Title
        tk.Label(self.toolbar, text="✦ WILDCARD EDITOR",
                 bg=COLORS["bg1"], fg=COLORS["accent"],
                 font=("Segoe UI", 11, "bold")).pack(side="left", padx=(12, 6))
        sep()

        self._btn(self.toolbar, "New",    self._new_file,    icon="⊞").pack(side="left")
        self._btn(self.toolbar, "Open",   self._open_file,   icon="📂").pack(side="left")
        self._btn(self.toolbar, "Save",   self._save_file,   fg=COLORS["accent"], icon="💾").pack(side="left")
        self._btn(self.toolbar, "Save As",self._save_file_as).pack(side="left")
        self._btn(self.toolbar, "Rename", self._rename_current_doc, fg=COLORS["accent2"], icon="✏").pack(side="left")
        sep()

        self.btn_back = self._btn(self.toolbar, "◀", self._nav_back, fg=COLORS["text2"])
        self.btn_back.pack(side="left")
        self.btn_fwd  = self._btn(self.toolbar, "▶", self._nav_forward, fg=COLORS["text2"])
        self.btn_fwd.pack(side="left")
        sep()

        self._btn(self.toolbar, "↩ Undo", self._do_undo, fg=COLORS["text1"]).pack(side="left")
        self._btn(self.toolbar, "↪ Redo", self._do_redo, fg=COLORS["text1"]).pack(side="left")
        sep()

        self._btn(self.toolbar, "Clone Lines", self._clone_lines, fg=COLORS["accent2"], icon="⧉").pack(side="left")
        self._btn(self.toolbar, "Wrap Wildcard", self._wrap_wildcard, fg=COLORS["accent3"], icon="⟨⟩").pack(side="left")
        sep()

        self._btn(self.toolbar, "Find", self._toggle_find, icon="🔍").pack(side="left")
        self._btn(self.toolbar, "Search All", self._show_search_all, fg=COLORS["accent2"], icon="⊞").pack(side="left")
        self._btn(self.toolbar, "⚕ Diagnose", self._show_diagnostics, fg=COLORS["warn"]).pack(side="left")
        self._btn(self.toolbar, "LoRA ±", self._show_lora_adjust, fg=COLORS["accent3"]).pack(side="left")
        sep()

        # Word Wrap toggle
        self.wrap_btn = tk.Label(self.toolbar, text="⏎ Wrap: ON",
                                  bg=COLORS["bg1"], fg=COLORS["accent3"],
                                  font=("Segoe UI", 10, "bold"),
                                  cursor="hand2", padx=9, pady=4)
        self.wrap_btn.pack(side="left")
        self.wrap_btn.bind("<Button-1>", lambda e: self._toggle_word_wrap())
        self.wrap_btn.bind("<Enter>", lambda e: self.wrap_btn.config(bg=COLORS["bg3"]))
        self.wrap_btn.bind("<Leave>", lambda e: self.wrap_btn.config(bg=COLORS["bg1"]))
        sep()

        self._btn(self.toolbar, "Reorganize", self._show_reorg_confirm, fg=COLORS["danger"], icon="⇄").pack(side="left")
        sep()
        self._btn(self.toolbar, "Hotkeys", self._show_hotkeys, icon="⌨").pack(side="left")
        self._btn(self.toolbar, "Settings", self._show_settings, icon="⚙").pack(side="left")
        sep()

        # Toggle tab bar visibility
        self.tab_bar_visible = True
        self.tab_toggle_btn = self._btn(self.toolbar, "Tabs ▾", self._toggle_tab_bar)
        self.tab_toggle_btn.pack(side="left")

        self._update_wrap_btn()

    def _build_main(self):
        # Outer pane: sidebar | editor
        self.paned = tk.PanedWindow(self.root, orient="horizontal",
                                     bg=COLORS["border2"], sashwidth=4,
                                     sashrelief="flat", bd=0)
        self.paned.pack(fill="both", expand=True)

        # ── SIDEBAR ──────────────────────────────────────────────────────
        self.sidebar = tk.Frame(self.paned, bg=COLORS["bg1"], width=self.cfg.get("sidebar_width", 280))
        self.sidebar.pack_propagate(False)
        self.paned.add(self.sidebar, minsize=160)

        # Inner vertical pane: folder tree | wc list
        self.side_paned = tk.PanedWindow(self.sidebar, orient="vertical",
                                          bg=COLORS["border2"], sashwidth=5,
                                          sashrelief="raised", bd=0)
        self.side_paned.pack(fill="both", expand=True)
        # Set sash to 2/3 folder, 1/3 wc list after geometry is known
        self.root.after(100, self._set_sidebar_sash)

        # ── Folder panel ──
        self.folder_frame = tk.Frame(self.side_paned, bg=COLORS["bg1"])
        self.side_paned.add(self.folder_frame, minsize=80)

        fhdr = tk.Frame(self.folder_frame, bg=COLORS["bg2"], height=30)
        fhdr.pack(fill="x")
        fhdr.pack_propagate(False)

        def small_btn(parent, text, cmd, fg=COLORS["text2"]):
            b = tk.Label(parent, text=text, bg=COLORS["bg2"], fg=fg,
                         font=("Segoe UI", 9), cursor="hand2", padx=6)
            b.pack(side="right", pady=3)
            b.bind("<Button-1>", lambda e: cmd())
            b.bind("<Enter>", lambda e: b.config(bg=COLORS["bg3"], fg=COLORS["text0"]))
            b.bind("<Leave>", lambda e: b.config(bg=COLORS["bg2"], fg=fg))
            return b

        small_btn(fhdr, "Sort ▾", self._show_sort_menu, fg=COLORS["text2"])
        small_btn(fhdr, "+ Folder", self._new_folder_dialog, fg=COLORS["accent"])
        small_btn(fhdr, "⬇ Import", self._import_folder_structure, fg=COLORS["accent3"])
        small_btn(fhdr, "✂ Isolated", self._remove_isolated_wildcards, fg=COLORS["danger"])

        tree_wrap = tk.Frame(self.folder_frame, bg=COLORS["bg1"])
        tree_wrap.pack(fill="both", expand=True)

        self.tree_scroll = ttk.Scrollbar(tree_wrap, orient="vertical")
        self.tree_scroll.pack(side="right", fill="y")

        self.tree = ttk.Treeview(tree_wrap, yscrollcommand=self.tree_scroll.set,
                                  selectmode="browse", show="tree")
        self.tree.pack(fill="both", expand=True)
        self.tree_scroll.config(command=self.tree.yview)
        self.tree.column("#0", minwidth=100)

        self.tree.tag_configure("folder",   foreground=COLORS["text0"],
                                 font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("doc",      foreground=COLORS["text1"],
                                 font=("Segoe UI", 10))
        self.tree.tag_configure("doc_active", foreground=COLORS["accent"],
                                 font=("Segoe UI", 10, "bold"))
        self.tree.tag_configure("unsorted_folder", foreground=COLORS["text2"],
                                 font=("Segoe UI", 10, "bold"))

        self.tree.bind("<ButtonRelease-1>", self._tree_click)
        self.tree.bind("<Double-Button-1>", self._tree_dbl_click)
        self.tree.bind("<Button-3>",        self._tree_right_click)
        self.tree.bind("<ButtonPress-1>",   self._drag_start)
        self.tree.bind("<B1-Motion>",       self._drag_motion)
        self.tree.bind("<<TreeviewOpen>>",  self._on_tree_open)
        self.tree.bind("<<TreeviewClose>>", self._on_tree_close)

        # ── WC List panel ──
        self.wc_frame = tk.Frame(self.side_paned, bg=COLORS["bg1"])
        self.side_paned.add(self.wc_frame, minsize=60)

        whdr = tk.Frame(self.wc_frame, bg=COLORS["bg2"], height=30)
        whdr.pack(fill="x")
        whdr.pack_propagate(False)
        tk.Label(whdr, text="WILDCARDS IN DOC", bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=8, pady=4)
        # Search for use button
        sfu_btn = tk.Label(whdr, text="Search Use", bg=COLORS["bg2"], fg=COLORS["accent2"],
                            font=("Segoe UI", 9), cursor="hand2", padx=6)
        sfu_btn.pack(side="right", pady=3)
        sfu_btn.bind("<Button-1>", lambda e: self._search_for_use())
        sfu_btn.bind("<Enter>", lambda e: sfu_btn.config(bg=COLORS["bg3"], fg=COLORS["text0"]))
        sfu_btn.bind("<Leave>", lambda e: sfu_btn.config(bg=COLORS["bg2"], fg=COLORS["accent2"]))
        # Open All button
        oa_btn = tk.Label(whdr, text="Open All", bg=COLORS["bg2"], fg=COLORS["accent3"],
                           font=("Segoe UI", 9), cursor="hand2", padx=6)
        oa_btn.pack(side="right", pady=3)
        oa_btn.bind("<Button-1>", lambda e: self._open_all_wildcards())
        oa_btn.bind("<Enter>", lambda e: oa_btn.config(bg=COLORS["bg3"], fg=COLORS["text0"]))
        oa_btn.bind("<Leave>", lambda e: oa_btn.config(bg=COLORS["bg2"], fg=COLORS["accent3"]))

        wc_wrap = tk.Frame(self.wc_frame, bg=COLORS["bg1"])
        wc_wrap.pack(fill="both", expand=True)

        self.wc_scroll = ttk.Scrollbar(wc_wrap, orient="vertical")
        self.wc_scroll.pack(side="right", fill="y")

        self.wc_list = tk.Listbox(wc_wrap, yscrollcommand=self.wc_scroll.set,
                                   bg=COLORS["bg1"], fg=COLORS["accent2"],
                                   selectbackground=COLORS["bg3"],
                                   selectforeground=COLORS["text0"],
                                   font=("Consolas", 10),
                                   relief="flat", bd=0,
                                   activestyle="none",
                                   highlightthickness=0)
        self.wc_list.pack(fill="both", expand=True)
        self.wc_scroll.config(command=self.wc_list.yview)
        self.wc_list.bind("<Button-1>",        self._wc_list_click)
        self.wc_list.bind("<Double-Button-1>", self._wc_list_dbl_click)
        self.wc_frame.bind("<MouseWheel>", lambda e: self.wc_list.yview_scroll(-1*(e.delta//120), "units"))
        self.folder_frame.bind("<MouseWheel>", lambda e: self.tree.yview_scroll(-1*(e.delta//120), "units"))

        # ── EDITOR AREA ───────────────────────────────────────────────────
        self.editor_frame = tk.Frame(self.paned, bg=COLORS["bg0"])
        self.paned.add(self.editor_frame, minsize=400)

        # Tab bar container (holds arrow + tab_bar + new btn)
        self.tab_bar_container = tk.Frame(self.editor_frame, bg=COLORS["bg1"], height=34)
        self.tab_bar_container.pack(fill="x")
        self.tab_bar_container.pack_propagate(False)

        # Left scroll arrow
        self._tab_scroll_left_btn = tk.Label(self.tab_bar_container, text="◀", bg=COLORS["bg1"],
                                              fg=COLORS["text2"], font=("Segoe UI", 10),
                                              cursor="hand2", padx=6)
        self._tab_scroll_left_btn.pack(side="left", fill="y")
        self._tab_scroll_left_btn.bind("<Button-1>", lambda e: self._scroll_tabs(-1))

        # Right scroll arrow
        self._tab_scroll_right_btn = tk.Label(self.tab_bar_container, text="▶", bg=COLORS["bg1"],
                                               fg=COLORS["text2"], font=("Segoe UI", 10),
                                               cursor="hand2", padx=6)
        self._tab_scroll_right_btn.pack(side="right", fill="y")
        self._tab_scroll_right_btn.bind("<Button-1>", lambda e: self._scroll_tabs(1))

        # Scrollable tab canvas
        self.tab_canvas = tk.Canvas(self.tab_bar_container, bg=COLORS["bg1"],
                                     height=34, highlightthickness=0, bd=0)
        self.tab_canvas.pack(side="left", fill="both", expand=True)

        # Frame inside canvas that holds actual tabs
        self.tab_bar = tk.Frame(self.tab_canvas, bg=COLORS["bg1"])
        self._tab_canvas_win = self.tab_canvas.create_window(0, 0, anchor="nw", window=self.tab_bar)
        self.tab_bar.bind("<Configure>", self._on_tabbar_configure)
        self._tab_offset = 0  # pixel scroll offset

        # Find & Replace panel (hidden initially)
        self.find_frame = tk.Frame(self.editor_frame, bg=COLORS["bg2"], relief="flat", bd=0)
        self._build_find_panel()

        # Editor with line numbers
        self.editor_wrap = tk.Frame(self.editor_frame, bg=COLORS["bg0"])
        self.editor_wrap.pack(fill="both", expand=True)

        # Line numbers — use a Canvas so we can pixel-perfectly position each number
        self.line_canvas = tk.Canvas(self.editor_wrap, width=48,
                                      bg=COLORS["bg1"], highlightthickness=0,
                                      bd=0, cursor="arrow")
        self.line_canvas.pack(side="left", fill="y")
        # Keep old name as alias so existing code that refs line_numbers still works
        self.line_numbers = self.line_canvas
        self.line_canvas.bind("<Button-1>", self._line_canvas_click)

        # Separator line
        tk.Frame(self.editor_wrap, bg=COLORS["border2"], width=1).pack(side="left", fill="y")

        # Text editor
        wrap_mode = "word" if self.word_wrap.get() else "none"
        self.editor = tk.Text(self.editor_wrap,
                               bg=COLORS["bg0"], fg=COLORS["text0"],
                               insertbackground=COLORS["accent"],
                               selectbackground=COLORS["sel_bg"],
                               selectforeground=COLORS["text0"],
                               font=(self.cfg.get("font_family","Consolas"), self.cfg["font_size"]),
                               relief="flat", bd=0, undo=True,
                               autoseparators=False,
                               maxundo=-1,
                               wrap=wrap_mode,
                               padx=12, pady=8,
                               highlightthickness=0,
                               tabs=("1c",))
        self.editor.pack(side="left", fill="both", expand=True)

        self.ed_scroll = ttk.Scrollbar(self.editor_wrap, orient="vertical",
                                        command=self.editor.yview)
        self.ed_scroll.pack(side="right", fill="y")

        def _on_yscroll(*args):
            """Called whenever the editor scrolls (mousewheel, scrollbar drag, keyboard)."""
            self.ed_scroll.set(*args)
            self._sync_scroll()

        self.editor.config(yscrollcommand=_on_yscroll)

        # Horizontal scrollbar (shown when wrap is off)
        self.ed_hscroll = ttk.Scrollbar(self.editor_frame, orient="horizontal",
                                         command=self.editor.xview)
        if not self.word_wrap.get():
            self.ed_hscroll.pack(fill="x")
        self.editor.config(xscrollcommand=self.ed_hscroll.set)

        # Editor tags
        self.editor.tag_configure("wildcard",
                                   background=COLORS["wc_hl"],
                                   foreground=COLORS["wc_hl_fg"],
                                   spacing3=0)
        # Scroll padding tag — visually invisible blank lines
        self.editor.tag_configure("_pad",
                                   foreground=COLORS["bg0"],
                                   background=COLORS["bg0"],
                                   selectforeground=COLORS["bg0"],
                                   selectbackground=COLORS["bg0"])
        self.editor.tag_configure("wc_active",
                                   background=COLORS["accent"],
                                   foreground=COLORS["bg0"])
        self.editor.tag_configure("find_hl",
                                   background="#1e3a1e",
                                   foreground=COLORS["accent3"])
        self.editor.tag_configure("find_cur",
                                   background=COLORS["warn"],
                                   foreground=COLORS["bg0"])
        self.editor.tag_configure("spell_err",
                                   foreground=COLORS["danger"],
                                   underline=True)
        self.editor.tag_configure("warn_tilde",
                                   background=COLORS["warn"],
                                   foreground=COLORS["bg0"])

        # Angle bracket highlight: <...> — teal tint (30% stronger than original)
        self.editor.tag_configure("hl_angle",
                                   background="#0d2524",
                                   foreground="#b8e8e8")

        # Parenthesis depth highlights — additive layers, stronger first level
        _paren_bgs = ["#1d2734","#243445","#2b4156","#324e67","#395b78","#406889","#47759a"]
        for i, bg in enumerate(_paren_bgs):
            self.editor.tag_configure(f"paren_d{i}", background=bg)

        # Curly brace depth highlights — same blue tint as parens (treated identically)
        for i, bg in enumerate(_paren_bgs):
            self.editor.tag_configure(f"curly_d{i}", background=bg)

        # Square bracket depth highlights — warm amber tint, stronger first level
        _sqbr_bgs  = ["#251914","#342016","#432718","#522e1a","#61351c","#703c1e","#7f4320"]
        for i, bg in enumerate(_sqbr_bgs):
            self.editor.tag_configure(f"sqbr_d{i}", background=bg)

        # Bold highlight for the cursor's matching bracket pair
        self.editor.tag_configure("bracket_match",
                                   font=(self.cfg.get("font_family","Consolas"),
                                         self.cfg.get("font_size",13), "bold"),
                                   foreground="#ffffff")

        # Ensure the built-in selection tag always renders on top of everything.
        # We reconfigure it with explicit colors and raise its priority.
        self.editor.tag_configure("sel",
                                   background=COLORS["sel_bg"],
                                   foreground=COLORS["text0"])
        # Raise sel above all custom tags so selection is always clearly visible
        self.editor.tag_raise("sel")

        self.editor.bind("<<Modified>>",       self._on_editor_modified)
        self.editor.bind("<KeyRelease>",        self._on_key_release)
        self.editor.bind("<ButtonRelease-1>",   self._on_editor_click)
        self.editor.bind("<Button-3>",          self._editor_right_click)
        self.editor.bind("<Double-Button-1>",   self._on_editor_dbl_click)
        self.editor.bind("<MouseWheel>",        self._on_editor_scroll)
        self.editor.bind("<Configure>",         lambda e: self._redraw_line_numbers())
        self.line_canvas.bind("<Configure>",    lambda e: self._redraw_line_numbers())
        # Undo / redo
        self.editor.bind("<Control-z>",         lambda e: (self._do_undo(), "break"))
        self.editor.bind("<Control-Z>",         lambda e: (self._do_undo(), "break"))
        self.editor.bind("<Control-y>",         lambda e: (self._do_redo(), "break"))
        self.editor.bind("<Control-Y>",         lambda e: (self._do_redo(), "break"))
        self.editor.bind("<Control-Shift-z>",   lambda e: (self._do_redo(), "break"))
        self.editor.bind("<Control-Shift-Z>",   lambda e: (self._do_redo(), "break"))
        # End key → jump to end of document
        self.editor.bind("<End>",               lambda e: self._on_end_key(e))
        # Mouse side buttons — Razer Viper Mini / gaming mice on Windows
        # Use the same multi-strategy approach proven in image_rater.py:
        # 1) Try Button-4 through Button-9
        # 2) XButton1 / XButton2 virtual events
        # 3) event_add virtual event mapping
        def _any_extra_btn(e, num):
            if num in (4, 8):  self._nav_back()
            elif num in (5, 9): self._nav_forward()
        for btn_num in (4, 5, 6, 7, 8, 9):
            try:
                self.root.bind(f"<Button-{btn_num}>",
                               lambda e, n=btn_num: _any_extra_btn(e, n))
            except Exception:
                pass
        for virt, cb in (("<XButton1>", lambda e: self._nav_back()),
                         ("<XButton2>", lambda e: self._nav_forward())):
            try:
                self.root.bind(virt, cb)
            except Exception:
                pass
        try:
            self.root.event_add("<<Mouse4>>", "<Button-8>")
            self.root.event_add("<<Mouse5>>", "<Button-9>")
            self.root.bind("<<Mouse4>>", lambda e: self._nav_back())
            self.root.bind("<<Mouse5>>", lambda e: self._nav_forward())
        except Exception:
            pass

    def _build_find_panel(self):
        fp = self.find_frame
        fp.columnconfigure(1, weight=1)

        # Row 1: Find
        tk.Label(fp, text="Find", bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9, "bold"), width=7).grid(row=0, column=0, padx=(8,4), pady=(8,2), sticky="w")

        self.find_var = tk.StringVar()
        self.find_entry = tk.Entry(fp, textvariable=self.find_var,
                                    bg=COLORS["bg3"], fg=COLORS["text0"],
                                    insertbackground=COLORS["accent"],
                                    relief="flat", bd=1,
                                    font=("Consolas", 10),
                                    highlightthickness=1,
                                    highlightcolor=COLORS["accent"],
                                    highlightbackground=COLORS["border2"])
        self.find_entry.grid(row=0, column=1, padx=4, pady=(8,2), sticky="ew")
        self.find_var.trace("w", lambda *a: self._do_find_highlight())

        mode_frame = tk.Frame(fp, bg=COLORS["bg2"])
        mode_frame.grid(row=0, column=2, padx=4, pady=(8,2))
        self._mode_btns = {}
        for m in ("Normal","Extended","Regex"):
            b = tk.Label(mode_frame, text=m, bg=COLORS["bg3"], fg=COLORS["text2"],
                         font=("Segoe UI", 9), cursor="hand2", padx=6, pady=2,
                         relief="flat")
            b.pack(side="left", padx=1)
            b.bind("<Button-1>", lambda e, mode=m.lower(): self._set_find_mode(mode))
            self._mode_btns[m.lower()] = b

        nav_frame = tk.Frame(fp, bg=COLORS["bg2"])
        nav_frame.grid(row=0, column=3, padx=4, pady=(8,2))
        def fbtn(parent, text, cmd):
            b = tk.Label(parent, text=text, bg=COLORS["bg3"], fg=COLORS["text1"],
                         font=("Segoe UI", 9, "bold"), cursor="hand2", padx=7, pady=2)
            b.pack(side="left", padx=1)
            b.bind("<Button-1>", lambda e: cmd())
            b.bind("<Enter>", lambda e: b.config(bg=COLORS["bg4"]))
            b.bind("<Leave>", lambda e: b.config(bg=COLORS["bg3"]))
            return b
        fbtn(nav_frame, "▲ Prev", self._find_prev)
        fbtn(nav_frame, "▼ Next", self._find_next)

        opt_frame = tk.Frame(fp, bg=COLORS["bg2"])
        opt_frame.grid(row=0, column=4, padx=4, pady=(8,2))
        self.find_case  = tk.BooleanVar()
        self.find_whole = tk.BooleanVar()
        self.find_wrap  = tk.BooleanVar(value=True)
        for var, text in [(self.find_case,"Case"),(self.find_whole,"Whole"),(self.find_wrap,"Wrap")]:
            cb = tk.Checkbutton(opt_frame, text=text, variable=var,
                                 bg=COLORS["bg2"], fg=COLORS["text1"],
                                 selectcolor=COLORS["bg3"], activebackground=COLORS["bg2"],
                                 font=("Segoe UI", 9),
                                 command=self._do_find_highlight)
            cb.pack(side="left", padx=2)

        self.find_status = tk.Label(fp, text="", bg=COLORS["bg2"], fg=COLORS["text2"],
                                     font=("Consolas", 9), width=10)
        self.find_status.grid(row=0, column=5, padx=4)

        close_btn = tk.Label(fp, text="✕", bg=COLORS["bg2"], fg=COLORS["text2"],
                              font=("Segoe UI", 11), cursor="hand2", padx=8)
        close_btn.grid(row=0, column=6, padx=(4,8), pady=(8,2))
        close_btn.bind("<Button-1>", lambda e: self._toggle_find())

        # Row 2: Replace
        tk.Label(fp, text="Replace", bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9, "bold"), width=7).grid(row=1, column=0, padx=(8,4), pady=(2,8), sticky="w")

        self.replace_var = tk.StringVar()
        repl_entry = tk.Entry(fp, textvariable=self.replace_var,
                               bg=COLORS["bg3"], fg=COLORS["text0"],
                               insertbackground=COLORS["accent"],
                               relief="flat", bd=1, font=("Consolas", 10),
                               highlightthickness=1,
                               highlightcolor=COLORS["accent"],
                               highlightbackground=COLORS["border2"])
        repl_entry.grid(row=1, column=1, padx=4, pady=(2,8), sticky="ew")

        repl_btn_frame = tk.Frame(fp, bg=COLORS["bg2"])
        repl_btn_frame.grid(row=1, column=2, columnspan=2, padx=4, pady=(2,8), sticky="w")
        fbtn(repl_btn_frame, "Replace", self._replace_current)
        fbtn(repl_btn_frame, "Replace All", self._replace_all)

        self.find_entry.bind("<Return>",       lambda e: self._find_next())
        self.find_entry.bind("<Shift-Return>", lambda e: self._find_prev())
        self.find_entry.bind("<Escape>",       lambda e: self._toggle_find())
        repl_entry.bind("<Return>",            lambda e: self._replace_current())
        repl_entry.bind("<Escape>",            lambda e: self._toggle_find())

        self._set_find_mode("normal")

    def _build_statusbar(self):
        self.statusbar = tk.Frame(self.root, bg=COLORS["bg1"], height=26)
        self.statusbar.pack(side="bottom", fill="x")
        self.statusbar.pack_propagate(False)

        def sb_label(text, fg=COLORS["text2"], **kw):
            l = tk.Label(self.statusbar, text=text, bg=COLORS["bg1"], fg=fg,
                         font=("Consolas", 9), padx=8, **kw)
            l.pack(side="left")
            return l

        self.sb_pos    = sb_label("Ln 1, Col 1")
        sb_label("│", fg=COLORS["border2"])
        self.sb_sel    = sb_label("No selection")
        sb_label("│", fg=COLORS["border2"])
        self.sb_words  = sb_label("Words: 0")
        sb_label("│", fg=COLORS["border2"])
        self.sb_file   = sb_label("Untitled")

        # Right side
        tk.Frame(self.statusbar, bg=COLORS["bg1"]).pack(side="left", fill="x", expand=True)
        self.sb_spell = tk.Label(self.statusbar,
                                  text="Spell ✓" if self.spell_enabled else "Spell ✗",
                                  bg=COLORS["bg1"],
                                  fg=COLORS["accent3"] if self.spell_enabled else COLORS["text2"],
                                  font=("Consolas", 9),
                                  padx=8, cursor="hand2")
        self.sb_spell.pack(side="right")
        self.sb_spell.bind("<Button-1>", lambda e: self._toggle_spell())

        sb_label("│", fg=COLORS["border2"]).pack(side="right")
        self.sb_wrap_str = sb_label(f"Wrap: {self.cfg['wrap_str']}", fg=COLORS["accent"])
        self.sb_wrap_str.pack(side="right")

    # ── TABS ─────────────────────────────────────────────────────────────────
    def _render_tabs(self):
        for widget in self.tab_bar.winfo_children():
            widget.destroy()

        # Pinned tabs first, then unpinned
        pinned   = [t for t in self.tabs if t in self._pinned_tabs]
        unpinned = [t for t in self.tabs if t not in self._pinned_tabs]
        ordered  = pinned + unpinned

        for tid in ordered:
            doc = self.tree_state["docs"].get(tid)
            if not doc:
                continue
            is_active = (tid == self.active_tab)
            modified  = doc.get("modified_unsaved", False)
            is_pinned = tid in self._pinned_tabs
            pin_prefix = "📌 " if is_pinned else ""
            name       = pin_prefix + doc["name"] + (" *" if modified else "")
            doc_color  = doc.get("color")

            if doc_color:
                tint     = _color_tint(doc_color, 0.10)
                tab_bg   = tint
                fg_color = COLORS["accent"] if is_active else doc_color
            else:
                tab_bg   = COLORS["bg2"] if is_active else COLORS["bg1"]
                fg_color = COLORS["accent"] if is_active else COLORS["text2"]

            frame = tk.Frame(self.tab_bar, bg=tab_bg, relief="flat", bd=0)
            frame.pack(side="left")

            if is_active:
                tk.Frame(frame, bg=COLORS["accent"], height=2).pack(fill="x")

            inner = tk.Frame(frame, bg=tab_bg)
            inner.pack(fill="both")

            lbl_font = ("Segoe UI", 10, "bold") if is_active else ("Segoe UI", 10)
            lbl = tk.Label(inner, text=name, bg=tab_bg,
                           fg=fg_color, font=lbl_font,
                           padx=10, pady=5, cursor="hand2")
            lbl.pack(side="left")

            close_fg = COLORS["text2"]
            close = tk.Label(inner, text="✕", bg=tab_bg,
                             fg=close_fg, font=("Segoe UI", 9),
                             padx=4, cursor="hand2")
            close.pack(side="right", pady=5)

            def make_binds(t_id, frm, lbl_w, cls_w, orig_bg, orig_fg):
                hover_bg = COLORS["bg3"]
                for w in (frm, lbl_w):
                    w.bind("<Button-1>", lambda e, i=t_id: self._switch_tab(i))
                    w.bind("<Button-3>", lambda e, i=t_id: self._tab_right_click(e, i))
                    w.bind("<Enter>",    lambda e, f=frm, lw=lbl_w, cw=cls_w: (
                        f.config(bg=hover_bg), lw.config(bg=hover_bg), cw.config(bg=hover_bg)))
                    w.bind("<Leave>",    lambda e, f=frm, lw=lbl_w, cw=cls_w, bg=orig_bg: (
                        f.config(bg=bg), lw.config(bg=bg), cw.config(bg=bg)))
                cls_w.bind("<Button-1>", lambda e, i=t_id: self._close_tab(i))
                cls_w.bind("<Enter>",    lambda e, cl=cls_w: cl.config(fg=COLORS["danger"]))
                cls_w.bind("<Leave>",    lambda e, cl=cls_w, f=close_fg: cl.config(fg=f))
            make_binds(tid, frame, lbl, close, tab_bg, fg_color)

        # New tab button at end
        new_btn = tk.Label(self.tab_bar, text=" ＋ ", bg=COLORS["bg1"],
                           fg=COLORS["text2"], font=("Segoe UI", 12),
                           cursor="hand2", pady=4)
        new_btn.pack(side="left")
        new_btn.bind("<Button-1>", lambda e: self._new_file())

        self.tab_bar.update_idletasks()
        self._scroll_active_tab_into_view()

    def _tab_right_click(self, event, doc_id):
        """Right-click context menu for a tab."""
        menu = tk.Menu(self.root, tearoff=0,
                       bg=COLORS["bg2"], fg=COLORS["text0"],
                       activebackground=COLORS["sel_bg"],
                       activeforeground=COLORS["text0"], relief="flat")

        is_pinned = doc_id in self._pinned_tabs
        pin_label = "Unpin Tab" if is_pinned else "📌 Pin Tab"

        menu.add_command(label=pin_label,
                         command=lambda: self._toggle_pin_tab(doc_id))
        menu.add_separator()
        menu.add_command(label="Search for Use",
                         command=lambda: (self._switch_tab(doc_id), self.root.after(50, self._search_for_use)))
        menu.add_command(label="Open All Contained Wildcards",
                         command=lambda: (self._switch_tab(doc_id), self.root.after(50, self._open_all_wildcards)))
        menu.add_separator()
        menu.add_command(label="Close",
                         command=lambda: self._close_tab(doc_id))
        menu.add_command(label="Close All Other Tabs",
                         command=lambda: self._close_all_other_tabs(doc_id))
        menu.add_separator()
        menu.add_command(label="Save",
                         command=lambda: self._save_file(doc_id=doc_id))

        menu.tk_popup(event.x_root, event.y_root)

    def _toggle_pin_tab(self, doc_id):
        if doc_id in self._pinned_tabs:
            self._pinned_tabs.discard(doc_id)
        else:
            self._pinned_tabs.add(doc_id)
        self._render_tabs()

    def _close_all_other_tabs(self, keep_id):
        """Close all tabs except keep_id and pinned tabs."""
        to_close = [t for t in list(self.tabs)
                    if t != keep_id and t not in self._pinned_tabs]
        for t in to_close:
            self._close_tab(t)

        # Scroll active tab into view
        self.tab_bar.update_idletasks()
        self._scroll_active_tab_into_view()

    def _on_tabbar_configure(self, event=None):
        self.tab_canvas.configure(scrollregion=self.tab_canvas.bbox("all"))

    def _scroll_tabs(self, direction):
        """Scroll the tab bar left (-1) or right (+1)."""
        canvas_w = self.tab_canvas.winfo_width()
        bar_w = self.tab_bar.winfo_reqwidth()
        max_offset = max(0, bar_w - canvas_w)
        self._tab_offset = max(0, min(max_offset, self._tab_offset + direction * 120))
        self.tab_canvas.xview_moveto(self._tab_offset / max(bar_w, 1))

    def _scroll_active_tab_into_view(self):
        """Ensure the active tab is visible in the canvas viewport."""
        if not self.active_tab: return
        try:
            for widget in self.tab_bar.winfo_children():
                if isinstance(widget, tk.Frame):
                    # Check if this frame's label matches active tab
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Frame):
                            for lbl in child.winfo_children():
                                if isinstance(lbl, tk.Label) and hasattr(lbl, '_tab_id'):
                                    pass
            # Simpler: just scroll to end if last tab is active
            self.tab_canvas.update_idletasks()
            bar_w = self.tab_bar.winfo_reqwidth()
            canvas_w = self.tab_canvas.winfo_width()
            if bar_w > canvas_w:
                # Only auto-scroll if active tab is last
                if self.tabs and self.tabs[-1] == self.active_tab:
                    self.tab_canvas.xview_moveto(1.0)
        except Exception:
            pass

    def _toggle_tab_bar(self):
        self.tab_bar_visible = not self.tab_bar_visible
        if self.tab_bar_visible:
            self.tab_bar_container.pack(fill="x", before=self.find_frame if self.find_frame.winfo_ismapped() else self.editor_wrap)
            self.tab_toggle_btn.config(text="Tabs ▾")
        else:
            self.tab_bar_container.pack_forget()
            self.tab_toggle_btn.config(text="Tabs ▸")

    def _open_tab(self, doc_id, push_nav=True):
        if doc_id not in self.tabs:
            self.tabs.append(doc_id)
        self._switch_tab(doc_id, push_nav=push_nav)

    def _switch_tab(self, doc_id, push_nav=True):
        if self.cfg.get("autosave") and self.active_tab:
            self._save_file(silent=True)

        # ── Flush any pending snapshot for the tab we're leaving ─────────────
        prev_id = self.active_tab
        if prev_id and prev_id != doc_id:
            # Sync the leaving tab's content to doc before anything else
            if prev_id in self.tree_state["docs"]:
                self.tree_state["docs"][prev_id]["content"] = self._get_real_content()
            self._save_undo_state(prev_id)   # flushes timer + saves cursor

        self.active_tab = doc_id
        doc = self.tree_state["docs"].get(doc_id)
        if doc:
            self.editor.config(state="normal")

            store    = self._undo_store.get(doc_id)
            has_hist = bool(store and store.get("stack"))

            # Choose content: top of existing stack, or doc's stored content
            if has_hist:
                content = store["stack"][store["pos"]][0]
            else:
                content = doc.get("content", "")

            # Load into editor without triggering undo machinery
            self._undo_inhibit = True
            self._undo_generation += 1
            self._load_count += 1
            try:
                self.editor.config(undo=False)
                self.editor.delete("1.0", "end")
                self.editor.insert("1.0", content)
                self.editor.config(undo=True)
                self.editor.edit_modified(False)
            finally:
                self._undo_inhibit = False
                self.root.after(0, self._dec_load_count)
            # NOTE: do NOT reset modified_unsaved here — it reflects real dirty state

            # Seed a fresh stack if this is the first visit
            if not has_hist:
                self._seed_undo_store(doc_id, content)

            # Restore cursor
            self._restore_undo_state(doc_id)

        # Clear find highlights — they belong to the previous tab's content
        self.editor.tag_remove("find_hl", "1.0", "end")
        self.editor.tag_remove("find_cur", "1.0", "end")
        self.find_matches = []
        self.find_current = -1
        # Re-run find after content and padding have fully settled
        if self.find_frame.winfo_ismapped():
            self.root.after(200, self._do_find_highlight)

        self._render_tabs()
        self._refresh_tree()
        self._update_line_numbers()
        self._update_status()
        self._update_wc_list()
        self._apply_wildcard_highlights()
        # Apply padding under inhibit so <<Modified>> from pad insert is ignored
        self._undo_inhibit = True
        self._undo_generation += 1
        self._load_count += 1
        try:
            self._apply_scroll_padding()
        finally:
            self._undo_inhibit = False
            self.root.after(0, self._dec_load_count)
        if self.spell_enabled:
            self.root.after(100, self._run_spell_check)
        if push_nav:
            self._nav_push(doc_id)
        if doc:
            self.sb_file.config(text=doc["name"] + (" — " + doc.get("path","") if doc.get("path") else ""))
            dirty = doc.get("modified_unsaved", False)
            prefix = "* " if dirty else ""
            self.root.title(f"{prefix}{doc['name']} — {APP_NAME}")
        else:
            self.root.title(APP_NAME)
        self._save_session()

    def _close_tab(self, doc_id):
        doc = self.tree_state["docs"].get(doc_id)
        # Pinned tab confirmation
        if doc_id in self._pinned_tabs:
            if not messagebox.askyesno("Close Pinned Tab",
                    f"'{doc['name']}' is pinned. Close it anyway?",
                    parent=self.root):
                return
            self._pinned_tabs.discard(doc_id)
        # Check if truly unsaved: dirty flag set AND content differs from saved hash
        truly_unsaved = False
        if doc and doc.get("modified_unsaved"):
            saved_hash = doc.get("saved_hash")
            content = self._get_real_content() if doc_id == self.active_tab else doc.get("content", "")
            truly_unsaved = not saved_hash or self._content_hash(content) != saved_hash
        if truly_unsaved:
            ans = messagebox.askyesnocancel("Unsaved Changes",
                f"Save '{doc['name']}' before closing?")
            if ans is None: return
            if ans: self._save_file(silent=False, doc_id=doc_id)
        if doc_id in self.tabs:
            idx = self.tabs.index(doc_id)
            self.tabs.remove(doc_id)
            self._cleanup_undo_store(doc_id)
            if self.active_tab == doc_id:
                nxt = self.tabs[idx] if idx < len(self.tabs) else (self.tabs[idx-1] if self.tabs else None)
                self.active_tab = nxt
                if nxt:
                    self._switch_tab(nxt, push_nav=False)
                else:
                    self.editor.delete("1.0", "end")
        self._render_tabs()
        self._refresh_tree()
        self._save_session()

    # ── FILE OPS ─────────────────────────────────────────────────────────────
    def _new_file(self):
        doc_id = str(uuid.uuid4())
        now = time.time()
        self.tree_state["docs"][doc_id] = {
            "id": doc_id, "name": "Untitled", "path": None,
            "content": "\n", "color": None,
            "created": now, "modified": now, "modified_unsaved": False,
            "saved_hash": self._content_hash("\n")
        }
        self.tree_state["unsorted"].append(doc_id)
        self.tabs.append(doc_id)
        self._switch_tab(doc_id)
        self._refresh_tree()
        save_tree_state(self.tree_state)
        # Place cursor at start of first real line after everything settles
        self.root.after(50, lambda: (
            self.editor.mark_set("insert", "1.0"),
            self.editor.see("1.0")))

    def _open_file(self):
        init = self.cfg.get("last_open_dir") or os.path.expanduser("~")
        paths = filedialog.askopenfilenames(
            title="Open File(s)",
            initialdir=init,
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
        for path in paths:
            self._load_file(path)
        if paths:
            self.cfg["last_open_dir"] = str(Path(paths[-1]).parent)
            save_config(self.cfg)

    def _load_file(self, path):
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")
            return
        doc_id = str(uuid.uuid4())
        st = os.stat(path)
        name = Path(path).stem
        self.tree_state["docs"][doc_id] = {
            "id": doc_id, "name": name, "path": path,
            "content": content, "color": None,
            "created": st.st_ctime, "modified": st.st_mtime,
            "modified_unsaved": False,
            "saved_hash": self._content_hash(content)
        }

        # Try to place into a matching explorer folder based on file path
        matched_folder = self._find_folder_for_path(path)
        if matched_folder is not None:
            matched_folder["docs"].append(doc_id)
        else:
            self.tree_state["unsorted"].append(doc_id)

        self.tabs.append(doc_id)
        self._switch_tab(doc_id)
        # _switch_tab inserts content which fires <<Modified>>, clear it again
        if doc_id in self.tree_state["docs"]:
            self.tree_state["docs"][doc_id]["modified_unsaved"] = False
        self.editor.edit_modified(False)
        self._refresh_tree()
        save_tree_state(self.tree_state)
        self._notify(f"Opened: {name}", "success")

    def _find_folder_for_path(self, file_path):
        """Given a file path, find the explorer folder whose reconstructed disk
        path best matches the file's parent directory. Returns the folder dict
        or None if no match found."""
        wc_dir = self.cfg.get("wc_dir", "")
        if not wc_dir or not os.path.isdir(wc_dir):
            return None

        file_dir = os.path.normpath(os.path.dirname(file_path))

        # Build folder_id -> absolute disk path using same logic as _do_reorganize
        folder_by_id = {f["id"]: f for f in self.tree_state["folders"]}
        path_cache = {}

        def get_folder_disk_path(folder_id):
            if folder_id in path_cache:
                return path_cache[folder_id]
            folder = folder_by_id.get(folder_id)
            if not folder:
                return None
            # Find parent
            parent = next(
                (f for f in self.tree_state["folders"]
                 if folder_id in f.get("children", [])),
                None
            )
            if parent is None:
                disk_path = os.path.join(wc_dir, folder["name"])
            else:
                parent_path = get_folder_disk_path(parent["id"])
                if parent_path is None:
                    return None
                disk_path = os.path.join(parent_path, folder["name"])
            path_cache[folder_id] = os.path.normpath(disk_path)
            return path_cache[folder_id]

        # Pre-compute all folder paths
        for f in self.tree_state["folders"]:
            get_folder_disk_path(f["id"])

        # Find exact match
        for folder_id, disk_path in path_cache.items():
            if disk_path == file_dir:
                return folder_by_id[folder_id]

        return None

    def _content_hash(self, text):
        """MD5 hash of content string for clean-state detection."""
        return hashlib.md5(text.encode("utf-8", errors="replace")).hexdigest()

    def _save_file(self, silent=False, doc_id=None):
        did = doc_id or self.active_tab
        if not did: return
        doc = self.tree_state["docs"].get(did)
        if not doc: return
        content = self._get_real_content() if did == self.active_tab else doc.get("content","")
        # Check for malformed wildcard wrappers before saving
        if did == self.active_tab and not silent:
            if self._check_wrapper_integrity(content):
                return  # user was warned and chose not to save
        doc["content"] = content
        doc["modified"] = time.time()
        doc["modified_unsaved"] = False
        doc["saved_hash"] = self._content_hash(content)  # record clean state
        if doc.get("path"):
            try:
                with open(doc["path"], "w", encoding="utf-8") as f:
                    f.write(content)
                if not silent: self._notify(f"Saved: {doc['name']}", "success")
            except Exception as e:
                messagebox.showerror("Save Error", str(e))
        else:
            self._save_file_as(doc_id=did)
            return
        self._render_tabs()
        self._refresh_tree()
        if did == self.active_tab:
            self.root.title(f"{doc['name']} — {APP_NAME}")
        save_tree_state(self.tree_state)

    def _save_file_as(self, doc_id=None):
        did = doc_id or self.active_tab
        if not did: return
        doc = self.tree_state["docs"].get(did)
        if not doc: return
        init_dir = self.cfg.get("last_open_dir") or self.cfg.get("wc_dir") or os.path.expanduser("~")
        path = filedialog.asksaveasfilename(
            title="Save As",
            initialdir=init_dir,
            initialfile=doc["name"] + ".txt",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")])
        if not path: return
        doc["path"] = path
        doc["name"] = Path(path).stem
        self.cfg["last_open_dir"] = str(Path(path).parent)
        save_config(self.cfg)
        self._save_file(doc_id=did)
        self._render_tabs()
        self._refresh_tree()

    def _dec_load_count(self):
        self._load_count = max(0, self._load_count - 1)

    # ── EDITOR EVENTS ────────────────────────────────────────────────────────
    def _on_editor_modified(self, event=None):
        if not self.editor.edit_modified(): return
        self.editor.edit_modified(False)
        # Ignore modifications from programmatic loads.
        # _undo_inhibit catches synchronous; _load_count catches the async
        # <<Modified>> that Tkinter queues even after edit_modified(False).
        if self._undo_inhibit or self._load_count > 0:
            return
        did = self.active_tab
        if did and did in self.tree_state["docs"]:
            doc = self.tree_state["docs"][did]
            doc["modified"] = time.time()
            if not doc.get("modified_unsaved"):
                doc["modified_unsaved"] = True
                self._render_tabs()
                self._refresh_tree()
                self.root.title(f"* {doc['name']} — {APP_NAME}")
        # Debounced snapshot
        if self._snap_timer is not None:
            self.root.after_cancel(self._snap_timer)
        self._snap_gen   = self._undo_generation
        self._snap_timer = self.root.after(500, self._commit_snapshot)
        # Debounce heavy UI work separately
        if self._highlight_job:
            self.root.after_cancel(self._highlight_job)
        self._highlight_job = self.root.after(600, self._deferred_update)
        if self._status_job:
            self.root.after_cancel(self._status_job)
        self._status_job = self.root.after(80, self._deferred_status)

    def _commit_snapshot(self):
        """Fires 500 ms after the last user-driven keystroke.
        Aborted if a snapshot was restored in the meantime (generation changed)."""
        self._snap_timer = None
        if self._undo_inhibit:
            return
        if self._undo_generation != self._snap_gen:
            # A restore happened after this timer was scheduled — skip push
            return
        self._push_undo_snapshot()

    def _check_clean_state(self):
        """After a debounced pause, check if content matches saved hash.
        If so, auto-clear the dirty flag (handles undo-back-to-saved)."""
        did = self.active_tab
        if not did or did not in self.tree_state["docs"]:
            return
        doc = self.tree_state["docs"][did]
        if not doc.get("modified_unsaved"):
            return  # already clean, nothing to do
        saved_hash = doc.get("saved_hash")
        if not saved_hash:
            return  # no reference hash (old in-memory doc), can't auto-clean
        current_content = self._get_real_content()
        if self._content_hash(current_content) == saved_hash:
            doc["modified_unsaved"] = False
            self._render_tabs()
            self._refresh_tree()
            self.root.title(f"{doc['name']} — {APP_NAME}")

    def _deferred_status(self):
        self._update_line_numbers()
        self._update_status()

    def _deferred_update(self):
        did = self.active_tab
        if did and did in self.tree_state["docs"]:
            self.tree_state["docs"][did]["content"] = self._get_real_content()
        self._apply_wildcard_highlights()
        self._update_wc_list()
        self._apply_scroll_padding()
        if self.spell_enabled:
            self._run_spell_check()
        self._check_clean_state()  # auto-clear dirty if content matches saved state
        save_tree_state(self.tree_state)

    def _on_key_release(self, event=None):
        self._update_status()
        self._update_bracket_highlights()

    def _on_end_key(self, event=None):
        """End key → jump to end of real document content (before padding)."""
        real_end = self._real_end_index()
        self.editor.mark_set("insert", real_end)
        self.editor.see(real_end)
        return "break"

    def _real_end_index(self):
        """Return the text index of the last real character before scroll padding."""
        try:
            ranges = self.editor.tag_ranges("_pad")
            if ranges and len(ranges) >= 2:
                pad_start = str(ranges[0])
                ln, col = map(int, pad_start.split("."))
                # Go to end of line before the pad's first newline
                return f"{ln}.0 - 1c" if ln > 1 else "1.end"
            return "end-1c"
        except Exception:
            return "end-1c"

    def _on_editor_click(self, event=None):
        self._update_status()
        self._update_bracket_highlights()

    def _on_editor_dbl_click(self, event=None):
        # Check if we double-clicked on a wildcard
        cursor = self.editor.index("insert")
        line_start = self.editor.index(f"{cursor} linestart")
        line_end   = self.editor.index(f"{cursor} lineend")
        line_text  = self.editor.get(line_start, line_end)
        click_col  = int(cursor.split(".")[1])
        wrap = re.escape(self.cfg["wrap_str"])
        pattern = wrap + r"([^\s]+?)" + wrap
        for m in re.finditer(pattern, line_text):
            if m.start() <= click_col <= m.end():
                self._jump_to_wildcard(m.group(1))
                return "break"

    # ── SCROLL-PAST-END PADDING ───────────────────────────────────────────────
    _PAD_LINES = 30  # visual lines of extra scroll space below last real line

    def _apply_scroll_padding(self):
        """Append invisible blank lines so the user can scroll 30 lines past the end.
        Padding operations are kept OUT of the undo stack by disabling undo briefly."""
        try:
            self._remove_scroll_padding()
            # Disable undo so pad insertion never appears in the undo stack
            self.editor.config(undo=False)
            pad = "\n" * self._PAD_LINES
            self.editor.insert("end", pad, "_pad")
            self.editor.tag_configure("_pad",
                foreground=COLORS["bg0"], background=COLORS["bg0"],
                selectforeground=COLORS["bg0"], selectbackground=COLORS["bg0"])
            self.editor.edit_modified(False)
            # Clamp cursor: if it somehow ended up inside the pad, bring it back
            try:
                pad_ranges = self.editor.tag_ranges("_pad")
                if pad_ranges and len(pad_ranges) >= 2:
                    pad_start = str(pad_ranges[0])
                    cursor    = self.editor.index("insert")
                    cs_ln, cs_col = map(int, pad_start.split("."))
                    ci_ln, ci_col = map(int, cursor.split("."))
                    if ci_ln > cs_ln or (ci_ln == cs_ln and ci_col >= cs_col):
                        real_end = f"{cs_ln - 1}.end" if cs_ln > 1 else "1.0"
                        self.editor.mark_set("insert", real_end)
            except Exception:
                pass
        except Exception:
            pass
        finally:
            try:
                self.editor.config(undo=True)
            except Exception:
                pass

    def _remove_scroll_padding(self):
        """Delete any _pad-tagged content from the editor widget."""
        try:
            ranges = self.editor.tag_ranges("_pad")
            if ranges and len(ranges) >= 2:
                self.editor.config(undo=False)
                self.editor.delete(ranges[0], ranges[-1])
                self.editor.edit_modified(False)
        except Exception:
            pass
        finally:
            try:
                self.editor.config(undo=True)
            except Exception:
                pass

    # ── PER-TAB UNDO HISTORY ─────────────────────────────────────────────────
    # Tkinter's Text undo stack cannot be transferred between widgets, so we
    # manage our own action-based snapshot stack per tab.
    #
    # _undo_store[doc_id] = {
    #   "stack": [(content, cursor), ...],   # undo history, oldest first
    #   "pos":   int,                         # current position in stack
    #   "cursor": "ln.col"                    # last known cursor position
    # }
    #
    # Every "action boundary" (space, return, paste, etc.) calls _push_undo_snapshot().
    # Ctrl+Z walks stack backward; Ctrl+Y walks forward.
    # On tab switch away we save cursor; on return we restore cursor.

    # ── PER-TAB UNDO / REDO ───────────────────────────────────────────────────
    #
    # Design
    # ------
    # _undo_store[doc_id] = {
    #     "stack": [(content_str, cursor_str), ...],   oldest → newest
    #     "pos"  : int,                                 index of current state
    # }
    #
    # Invariant: stack is never empty once seeded; pos is always a valid index.
    #
    # Snapshots are pushed by _commit_snapshot(), which is called 500 ms after
    # the last <<Modified>> event.  Programmatic edits (replace, LoRA adjust)
    # call _push_undo_snapshot() directly — with _undo_inhibit=True during the
    # edit so the Modified handler doesn't schedule a duplicate timer.
    #
    # Ctrl+Z  → _do_undo()  : pos -= 1, apply stack[pos]
    # Ctrl+Y / Ctrl+Shift+Z → _do_redo() : pos += 1, apply stack[pos]
    #
    # Tab-away : flush any pending snap timer immediately, save cursor.
    # Tab-back : load stack[pos] into editor, restore cursor.
    # -------------------------------------------------------------------------

    def _get_or_create_undo_widget(self, doc_id):
        """Compat shim used by a few call sites; just ensures the store exists."""
        return self._undo_store.setdefault(doc_id, {"stack": [], "pos": 0})

    def _seed_undo_store(self, doc_id, content):
        """Initialise a fresh undo stack for doc_id with content as the base state.
        Called only when the store doesn't exist yet (first open of a tab)."""
        self._undo_store[doc_id] = {
            "stack": [(content, "1.0")],
            "pos":   0,
        }

    def _push_undo_snapshot(self, doc_id=None):
        """Append the current editor state to the undo stack.
        Safe to call at any time; no-ops if inhibited or content unchanged."""
        if self._undo_inhibit:
            return
        did = doc_id or self.active_tab
        if not did:
            return

        # Read authoritative content
        if did == self.active_tab:
            content = self._get_real_content()
            try:
                cursor = self.editor.index("insert")
            except Exception:
                cursor = "1.0"
        else:
            content = self.tree_state["docs"].get(did, {}).get("content", "")
            cursor  = self._undo_store.get(did, {}).get("stack",
                      [("", "1.0")])[self._undo_store.get(did, {}).get("pos", 0)][1]

        store = self._undo_store.get(did)
        if not store:
            self._seed_undo_store(did, content)
            return

        stack = store["stack"]
        pos   = store["pos"]

        # If content unchanged, just update the cursor on the current entry
        if stack[pos][0] == content:
            stack[pos] = (content, cursor)
            return

        # Discard any redo history above current position
        del stack[pos + 1:]

        stack.append((content, cursor))
        store["pos"] = len(stack) - 1

        # Cap memory usage
        if len(stack) > 300:
            stack[:] = stack[-300:]
            store["pos"] = len(stack) - 1

    def _flush_snap_timer(self):
        """If a debounced snapshot is pending, execute it immediately."""
        if self._snap_timer is not None:
            self.root.after_cancel(self._snap_timer)
            self._snap_timer = None
            if not self._undo_inhibit and self._undo_generation == self._snap_gen:
                self._push_undo_snapshot()

    def _save_undo_state(self, doc_id):
        """Called on tab-away: flush any pending snapshot and record cursor."""
        if not doc_id:
            return
        self._flush_snap_timer()
        store = self._undo_store.get(doc_id)
        if store:
            try:
                pos = store["pos"]
                content, _ = store["stack"][pos]
                store["stack"][pos] = (content, self.editor.index("insert"))
            except Exception:
                pass

    def _restore_undo_state(self, doc_id):
        """Called on tab-return: place cursor at last known position."""
        store = self._undo_store.get(doc_id)
        if not store:
            return
        try:
            _, cursor = store["stack"][store["pos"]]
            self.editor.mark_set("insert", cursor)
            self.editor.see(cursor)
        except Exception:
            pass

    def _cleanup_undo_store(self, doc_id):
        self._undo_store.pop(doc_id, None)

    def _do_undo(self):
        did = self.active_tab
        if not did:
            return
        # Cancel any pending snapshot timer — we're about to move through history,
        # we don't want a stale timer pushing the current content afterwards.
        # But first capture the current content if it differs from stack top,
        # so the user can redo back to what they had.
        self._flush_snap_timer()
        store = self._undo_store.get(did)
        if not store or store["pos"] <= 0:
            return
        store["pos"] -= 1
        content, cursor = store["stack"][store["pos"]]
        self._apply_snapshot(content, cursor)

    def _do_redo(self):
        did = self.active_tab
        if not did:
            return
        # Cancel any pending timer — redo should not be confused by a queued snapshot
        if self._snap_timer is not None:
            self.root.after_cancel(self._snap_timer)
            self._snap_timer = None
        store = self._undo_store.get(did)
        if not store or store["pos"] >= len(store["stack"]) - 1:
            return
        store["pos"] += 1
        content, cursor = store["stack"][store["pos"]]
        self._apply_snapshot(content, cursor)

    def _apply_snapshot(self, content, cursor):
        """Load content into the editor without touching the undo stack."""
        self._undo_generation += 1
        self._undo_inhibit = True
        self._load_count += 1
        try:
            self._remove_scroll_padding()
            self.editor.config(undo=False)
            self.editor.delete("1.0", "end")
            self.editor.insert("1.0", content)
            self.editor.config(undo=True)
            self.editor.edit_modified(False)
            self._apply_scroll_padding()
            try:
                self.editor.mark_set("insert", cursor)
                self.editor.see(cursor)
            except Exception:
                pass
            did = self.active_tab
            if did and did in self.tree_state["docs"]:
                self.tree_state["docs"][did]["content"] = content
                # Update dirty/clean indicator
                doc = self.tree_state["docs"][did]
                saved_hash = doc.get("saved_hash")
                if saved_hash:
                    is_clean = self._content_hash(content) == saved_hash
                    if is_clean and doc.get("modified_unsaved"):
                        doc["modified_unsaved"] = False
                        self._render_tabs()
                        self._refresh_tree()
                        self.root.title(f"{doc['name']} — {APP_NAME}")
                    elif not is_clean and not doc.get("modified_unsaved"):
                        doc["modified_unsaved"] = True
                        self._render_tabs()
                        self._refresh_tree()
                        self.root.title(f"* {doc['name']} — {APP_NAME}")
        finally:
            self._undo_inhibit = False
            self.root.after(0, self._dec_load_count)
        self._update_line_numbers()
        self._update_status()
        self.root.after(50, self._apply_wildcard_highlights)

    def _sync_store_to_editor(self, store):
        """Compat shim — not used."""
        pass

    def _get_real_content(self):
        """Return editor content excluding scroll-padding lines."""
        try:
            ranges = self.editor.tag_ranges("_pad")
            if ranges and len(ranges) >= 2:
                # Get everything before the first pad character
                raw = self.editor.get("1.0", ranges[0])
                # Strip any trailing newline that was the boundary between real content
                # and the first pad newline (there will be exactly one)
                if raw.endswith("\n"):
                    raw = raw[:-1]
                return raw
            return self.editor.get("1.0", "end-1c")
        except Exception:
            return self.editor.get("1.0", "end-1c")

    def _sync_scroll(self, event=None):
        """Redraw line numbers to match current editor scroll position."""
        self._redraw_line_numbers()

    def _on_editor_scroll(self, event):
        """Handle MouseWheel: Ctrl+scroll = font size, else normal scroll."""
        if event.state & 0x4:
            delta = 1 if event.delta > 0 else -1
            new_size = max(8, min(32, self.cfg["font_size"] + delta))
            if new_size != self.cfg["font_size"]:
                self.cfg["font_size"] = new_size
                fam = self.cfg.get("font_family", "Consolas")
                self.editor.config(font=(fam, new_size))
                self._apply_scroll_padding()
                self._redraw_line_numbers()
                save_config(self.cfg)
            return "break"
        units = -1 * (event.delta // 120) * 2  # ×2 scroll speed
        self.editor.yview_scroll(units, "units")
        self._redraw_line_numbers()
        return "break"


    def _update_line_numbers(self):
        self._redraw_line_numbers()

    def _redraw_line_numbers(self):
        """Draw line numbers on canvas using dlineinfo — pixel-perfect, visible lines only.
        dlineinfo(index) returns (x, y, w, h, baseline) with y relative to widget top,
        which equals the canvas y since canvas sits flush against the editor."""
        canvas = self.line_canvas
        canvas.delete("all")
        try:
            fam      = self.cfg.get("font_family", "Consolas")
            fs       = self.cfg.get("font_size", 13)
            canvas_h = canvas.winfo_height()
            if canvas_h < 2:
                return

            # Find where real content ends (before pad lines)
            pad_ranges = self.editor.tag_ranges("_pad")
            if pad_ranges and len(pad_ranges) >= 2:
                # The pad starts at this index — only draw numbers for lines before it
                pad_start_line = int(str(pad_ranges[0]).split(".")[0])
            else:
                pad_start_line = None  # no padding, draw all lines

            content   = self._get_real_content()
            n_logical = content.count("\n") + 1

            for ln in range(1, n_logical + 1):
                # Skip any line that is in or beyond the pad zone
                if pad_start_line is not None and ln >= pad_start_line:
                    break
                info = self.editor.dlineinfo(f"{ln}.0")
                if info is None:
                    # Line is scrolled off-screen — not visible, skip
                    continue
                x, y, w, h, baseline = info
                if y >= canvas_h:
                    break   # below visible area
                if y + h <= 0:
                    continue  # above visible area (shouldn't happen but be safe)
                # Draw number right-aligned at x=44, vertically centred in the line
                canvas.create_text(
                    44, y + h // 2,
                    text=str(ln), anchor="e",
                    fill=COLORS["text2"],
                    font=(fam, fs)
                )
        except Exception:
            pass

    def _line_canvas_click(self, event):
        """Click on a line number to select the entire text of that line."""
        try:
            # Use dlineinfo to find which logical line was clicked
            fam      = self.cfg.get("font_family", "Consolas")
            fs       = self.cfg.get("font_size", 13)
            content  = self._get_real_content()
            n_logical = content.count("\n") + 1
            pad_ranges = self.editor.tag_ranges("_pad")
            pad_start_line = int(str(pad_ranges[0]).split(".")[0]) if pad_ranges and len(pad_ranges) >= 2 else None

            clicked_y = event.y
            best_ln   = None
            best_dist = float("inf")

            for ln in range(1, n_logical + 1):
                if pad_start_line and ln >= pad_start_line:
                    break
                info = self.editor.dlineinfo(f"{ln}.0")
                if info is None:
                    continue
                x, y, w, h, baseline = info
                if y > clicked_y + h:
                    break
                # Check if the click falls within this line's y range
                if y <= clicked_y <= y + h:
                    best_ln = ln
                    break
                # Track nearest as fallback
                dist = abs(clicked_y - (y + h // 2))
                if dist < best_dist:
                    best_dist = dist
                    best_ln = ln

            if best_ln is None:
                return

            # Select from start of line to end of line (including the newline so
            # the selection visually covers the whole row)
            line_start = f"{best_ln}.0"
            line_end   = f"{best_ln}.end"
            self.editor.tag_remove("sel", "1.0", "end")
            self.editor.tag_add("sel", line_start, line_end)
            self.editor.mark_set("insert", line_start)
            self.editor.see(line_start)
            self.editor.focus_set()
        except Exception:
            pass

    def _update_status(self):
        try:
            idx = self.editor.index("insert")
            ln, col = idx.split(".")
            self.sb_pos.config(text=f"Ln {ln}, Col {int(col)+1}")
            try:
                sel_start = self.editor.index("sel.first")
                sel_end   = self.editor.index("sel.last")
                sel_text  = self.editor.get(sel_start, sel_end)
                self.sb_sel.config(text=f"{len(sel_text)} chars")
            except tk.TclError:
                self.sb_sel.config(text="No sel")
            content = self._get_real_content()
            words = len(content.split()) if content.strip() else 0
            self.sb_words.config(text=f"Words: {words}")
        except Exception:
            pass

    # ── WILDCARD HIGHLIGHTING ────────────────────────────────────────────────
    def _apply_wildcard_highlights(self):
        self.editor.tag_remove("wildcard", "1.0", "end")
        wrap = self.cfg["wrap_str"]
        esc = re.escape(wrap)
        pattern = esc + r"([^\s]+?)" + esc
        content = self._get_real_content()
        for m in re.finditer(pattern, content):
            start_idx = f"1.0+{m.start()}c"
            end_idx   = f"1.0+{m.end()}c"
            self.editor.tag_add("wildcard", start_idx, end_idx)
        # Skip bracket highlights for very large content to keep UI responsive
        if len(content) < 50_000:
            self._apply_bracket_highlights()
            self._update_bracket_highlights()
        else:
            # Still clear stale bracket tags so nothing wrong shows
            for i in range(7):
                self.editor.tag_remove(f"paren_d{i}", "1.0", "end")
                self.editor.tag_remove(f"sqbr_d{i}",  "1.0", "end")
                self.editor.tag_remove(f"curly_d{i}", "1.0", "end")
            self.editor.tag_remove("hl_angle", "1.0", "end")

    def _apply_bracket_highlights(self):
        """Apply static depth-based highlights for (...), [...], and <...> spans.
        Called when content changes. Cursor-based bold is handled separately."""
        content = self._get_real_content()
        ed = self.editor

        # Clear all static bracket tags
        for i in range(7):
            ed.tag_remove(f"paren_d{i}", "1.0", "end")
            ed.tag_remove(f"sqbr_d{i}",  "1.0", "end")
            ed.tag_remove(f"curly_d{i}", "1.0", "end")
        ed.tag_remove("hl_angle", "1.0", "end")

        # ── Angle brackets <...> ─────────────────────────────────────────────
        # Simple non-nested scan
        for m in re.finditer(r"<[^<>]*>", content):
            s = f"1.0+{m.start()}c"
            e = f"1.0+{m.end()}c"
            ed.tag_add("hl_angle", s, e)

        # ── Parentheses (...) — depth tracking ───────────────────────────────
        # For each character, track nesting depth. Apply the *span* of each
        # matched pair at depth d (capped to 6) using tag paren_d{d}.
        # We use a stack to find each matching pair, then tag the whole span.
        self._apply_depth_tags(content, "(", ")", "paren_d")

        # ── Square brackets [...] ─────────────────────────────────────────────
        self._apply_depth_tags(content, "[", "]", "sqbr_d")

        # ── Curly braces {...} — same depth colours as parentheses ───────────
        self._apply_depth_tags(content, "{", "}", "curly_d")

        # Raise wildcard tag so it stays on top of bracket bg colours
        try:
            ed.tag_raise("wildcard")
            ed.tag_raise("spell_err")
            ed.tag_raise("find_hl")
            ed.tag_raise("find_cur")
            ed.tag_raise("sel")   # selection always on top
        except Exception:
            pass

    def _apply_depth_tags(self, content, open_ch, close_ch, tag_prefix):
        """Tag matched bracket/paren spans with depth-based background tags.
        Batches tag_add calls into runs to minimise widget round-trips."""
        ed = self.editor

        # Build depth map in pure Python — fast, no widget calls
        depth_at = [0] * len(content)
        d = 0
        for i, ch in enumerate(content):
            if ch == open_ch:
                d += 1
            depth_at[i] = d
            if ch == close_ch and d > 0:
                d -= 1

        # Apply tags level by level using contiguous run detection
        # Collect all (start,end) pairs first, then do tag_add in one shot per run
        for level in range(7):
            threshold = level + 1
            runs = []
            in_run = False
            run_start = 0
            for i in range(len(depth_at)):
                meets = depth_at[i] >= threshold
                if meets and not in_run:
                    run_start = i
                    in_run = True
                elif not meets and in_run:
                    runs.append((run_start, i))
                    in_run = False
            if in_run:
                runs.append((run_start, len(depth_at)))

            if not runs:
                continue
            # Batch: build one long arg list and call tag_add once per run
            tag = f"{tag_prefix}{level}"
            for rs, re_ in runs:
                ed.tag_add(tag, f"1.0+{rs}c", f"1.0+{re_}c")

    def _update_bracket_highlights(self):
        """Update the cursor-position bold highlights for matching bracket/paren pairs.
        Called on every key/click event — fast, only touches bracket_match tag."""
        ed = self.editor
        ed.tag_remove("bracket_match", "1.0", "end")
        try:
            cursor = ed.index("insert")
            cur_line = int(cursor.split(".")[0])
            cur_col  = int(cursor.split(".")[1])
            content  = self._get_real_content()

            # Convert cursor to linear offset in content
            lines = content.split("\n")
            offset = sum(len(lines[i]) + 1 for i in range(cur_line - 1)) + cur_col
            # Clamp to valid range
            offset = max(0, min(offset, len(content)))

            # Check character at cursor and one before (bracket could be just typed)
            PAIRS = {"(": ")", "[": "]", "{": "}", ")": "(", "]": "[", "}": "{"}
            ch_at   = content[offset]     if offset < len(content) else ""
            ch_prev = content[offset - 1] if offset > 0 else ""

            # Prefer the character AT the cursor; fall back to the one before
            if ch_at in PAIRS:
                anchor = offset
                anchor_ch = ch_at
            elif ch_prev in PAIRS:
                anchor = offset - 1
                anchor_ch = ch_prev
            else:
                return  # cursor not near a bracket

            # Find matching bracket
            if anchor_ch in ("(", "[", "{"):
                close_ch = PAIRS[anchor_ch]
                depth = 0
                for i in range(anchor, len(content)):
                    c = content[i]
                    if c == anchor_ch:
                        depth += 1
                    elif c == close_ch:
                        depth -= 1
                        if depth == 0:
                            match = i
                            break
                else:
                    return  # no match
            else:  # closing bracket
                open_ch = PAIRS[anchor_ch]
                depth = 0
                for i in range(anchor, -1, -1):
                    c = content[i]
                    if c == anchor_ch:
                        depth += 1
                    elif c == open_ch:
                        depth -= 1
                        if depth == 0:
                            match = i
                            break
                else:
                    return  # no match

            # Convert both positions back to text indices
            def offset_to_idx(off):
                row = 1
                col = 0
                for ln in lines:
                    ln_len = len(ln) + 1  # +1 for newline
                    if off < ln_len:
                        col = off
                        break
                    off -= ln_len
                    row += 1
                return f"{row}.{col}"

            anchor_idx = offset_to_idx(anchor)
            match_idx  = offset_to_idx(match)

            ed.tag_add("bracket_match", anchor_idx, f"{anchor_idx}+1c")
            ed.tag_add("bracket_match", match_idx,  f"{match_idx}+1c")
            ed.tag_raise("bracket_match")
            try:
                ed.tag_raise("sel")
            except Exception:
                pass
        except Exception:
            pass

    def _highlight_wc_in_editor(self, name):
        """Highlight all instances of a specific wildcard name."""
        self.editor.tag_remove("wc_active", "1.0", "end")
        wrap = self.cfg["wrap_str"]
        target = wrap + name + wrap
        start = "1.0"
        first = None
        while True:
            pos = self.editor.search(target, start, stopindex="end", nocase=False)
            if not pos: break
            end = f"{pos}+{len(target)}c"
            self.editor.tag_add("wc_active", pos, end)
            if first is None: first = pos
            start = end
        if first:
            self.editor.see(first)

    # ── WILDCARD NAVIGATION ──────────────────────────────────────────────────
    def _jump_to_wildcard(self, name):
        """Jump to a doc matching wildcard name, searching disk subfolders too."""
        # Search open docs first
        found_id = None
        for did, doc in self.tree_state["docs"].items():
            if doc["name"].lower() == name.lower():
                found_id = did
                break
        if found_id:
            self._open_tab(found_id)
            self._notify(f"Jumped to: {name}", "info")
            return
        # Search disk
        wc_dir = self.cfg.get("wc_dir", "")
        if wc_dir and os.path.isdir(wc_dir):
            for root, dirs, files in os.walk(wc_dir):
                for f in files:
                    if Path(f).stem.lower() == name.lower() and f.endswith(".txt"):
                        fp = os.path.join(root, f)
                        self._load_file(fp)
                        self._notify(f"Opened from disk: {name}", "success")
                        return
        # Create new
        self._notify(f'Wildcard "{name}" not found — creating…', "warn")
        doc_id = str(uuid.uuid4())
        now = time.time()
        self.tree_state["docs"][doc_id] = {
            "id": doc_id, "name": name, "path": None,
            "content": "", "color": None,
            "created": now, "modified": now, "modified_unsaved": False
        }
        # Add to "new wildcards" folder
        nwf = next((f for f in self.tree_state["folders"] if f["name"].lower() == "new wildcards"), None)
        if not nwf:
            nwf = {"id": str(uuid.uuid4()), "name": "new wildcards",
                   "color": COLORS["accent2"], "open": True,
                   "children": [], "docs": []}
            self.tree_state["folders"].append(nwf)
        nwf["docs"].append(doc_id)
        self.tabs.append(doc_id)
        self._switch_tab(doc_id)
        self._refresh_tree()
        save_tree_state(self.tree_state)

    # ── WILDCARD LIST (sidebar) ──────────────────────────────────────────────
    def _update_wc_list(self):
        self.wc_list.delete(0, "end")
        wrap = self.cfg["wrap_str"]
        esc = re.escape(wrap)
        pattern = esc + r"([^\s]+?)" + esc
        content = self._get_real_content()
        found = sorted(set(m.group(1) for m in re.finditer(pattern, content)))
        for name in found:
            self.wc_list.insert("end", f" ❯ {wrap}{name}{wrap}")
        self._wc_names = found  # cache for click handler

    def _wc_list_click(self, event):
        idx = self.wc_list.nearest(event.y)
        if idx < 0 or idx >= len(getattr(self, "_wc_names", [])):
            return
        name = self._wc_names[idx]
        self._highlight_wc_in_editor(name)

    def _wc_list_dbl_click(self, event):
        idx = self.wc_list.nearest(event.y)
        if idx < 0 or idx >= len(getattr(self, "_wc_names", [])):
            return
        name = self._wc_names[idx]
        self._jump_to_wildcard(name)

    def _open_all_wildcards(self):
        """Open every wildcard found in the current document."""
        names = getattr(self, "_wc_names", [])
        if not names:
            self._notify("No wildcards found in this document", "warn")
            return
        opened = 0
        for name in names:
            # Check already open docs
            found_id = next((did for did, d in self.tree_state["docs"].items()
                             if d["name"].lower() == name.lower()), None)
            if found_id:
                if found_id not in self.tabs:
                    self.tabs.append(found_id)
                opened += 1
                continue
            # Search disk
            wc_dir = self.cfg.get("wc_dir", "")
            disk_path = None
            if wc_dir and os.path.isdir(wc_dir):
                for root_dir, dirs, files in os.walk(wc_dir):
                    for f in files:
                        if Path(f).stem.lower() == name.lower() and f.endswith(".txt"):
                            disk_path = os.path.join(root_dir, f)
                            break
                    if disk_path:
                        break
            if disk_path:
                # Load silently without switching tab
                try:
                    with open(disk_path, "r", encoding="utf-8", errors="replace") as f:
                        content = f.read()
                    doc_id = str(uuid.uuid4())
                    st = os.stat(disk_path)
                    self.tree_state["docs"][doc_id] = {
                        "id": doc_id, "name": Path(disk_path).stem,
                        "path": disk_path, "content": content, "color": None,
                        "created": st.st_ctime, "modified": st.st_mtime,
                        "modified_unsaved": False
                    }
                    self.tree_state["unsorted"].append(doc_id)
                    self.tabs.append(doc_id)
                    opened += 1
                except Exception:
                    pass
            else:
                # Create empty new wildcard doc
                doc_id = str(uuid.uuid4())
                now = time.time()
                self.tree_state["docs"][doc_id] = {
                    "id": doc_id, "name": name, "path": None,
                    "content": "", "color": None,
                    "created": now, "modified": now, "modified_unsaved": False
                }
                nwf = next((f for f in self.tree_state["folders"]
                            if f["name"].lower() == "new wildcards"), None)
                if not nwf:
                    nwf = {"id": str(uuid.uuid4()), "name": "new wildcards",
                           "color": COLORS["accent2"], "open": True,
                           "children": [], "docs": []}
                    self.tree_state["folders"].append(nwf)
                nwf["docs"].append(doc_id)
                self.tabs.append(doc_id)
                opened += 1

        self._render_tabs()
        self._refresh_tree()
        save_tree_state(self.tree_state)
        self._notify(f"Opened {opened} wildcard(s)", "success")

    # ── SEARCH FOR USE ────────────────────────────────────────────────────────
    def _search_for_use(self):
        """Search the wildcards folder for any .txt file that uses the current
        document's name as a wildcard reference (e.g. __docname__)."""
        if not self.active_tab:
            self._notify("No document is currently open.", "warn")
            return
        doc = self.tree_state["docs"].get(self.active_tab)
        if not doc:
            return
        name   = doc["name"]
        wc_dir = self.cfg.get("wc_dir", "")
        wrap   = self.cfg.get("wrap_str", "__")
        target = wrap + name + wrap

        if not wc_dir or not os.path.isdir(wc_dir):
            self._notify("Wildcards directory not set or missing.", "warn")
            return

        matches = []   # list of (filepath, line_number, line_text)
        try:
            for root_dir, dirs, files in os.walk(wc_dir):
                for fname in sorted(files):
                    if not fname.endswith(".txt"):
                        continue
                    fpath = os.path.join(root_dir, fname)
                    try:
                        with open(fpath, "r", encoding="utf-8", errors="replace") as f:
                            for lineno, line in enumerate(f, 1):
                                if target in line:
                                    matches.append((fpath, lineno, line.rstrip()))
                    except Exception:
                        pass
        except Exception as e:
            self._notify(f"Search error: {e}", "warn")
            return

        if not matches:
            messagebox.showinfo("Search for Use",
                f'No files reference {target}', parent=self.root)
            return

        # ── Show results window ───────────────────────────────────────────────
        win = tk.Toplevel(self.root)
        win.title(f"Uses of {target}")
        win.configure(bg=COLORS["bg2"])
        win.resizable(True, True)
        x = self.root.winfo_x() + 60
        y = self.root.winfo_y() + 60
        win.geometry(f"620x420+{x}+{y}")

        tk.Label(win, text=f'Files referencing  {target}',
                 bg=COLORS["bg2"], fg=COLORS["accent2"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(12,2))
        tk.Label(win, text=f'{len(matches)} occurrence(s) in {len(set(m[0] for m in matches))} file(s)',
                 bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9)).pack(anchor="w", padx=14, pady=(0,8))

        list_frame = tk.Frame(win, bg=COLORS["bg1"])
        list_frame.pack(fill="both", expand=True, padx=14, pady=(0,8))
        sb = ttk.Scrollbar(list_frame, orient="vertical")
        sb.pack(side="right", fill="y")
        lb = tk.Listbox(list_frame, bg=COLORS["bg1"], fg=COLORS["text0"],
                         selectbackground=COLORS["sel_bg"],
                         font=("Consolas", 10), relief="flat", bd=0,
                         activestyle="none", yscrollcommand=sb.set)
        lb.pack(fill="both", expand=True)
        sb.config(command=lb.yview)

        for fpath, lineno, line in matches:
            rel = os.path.relpath(fpath, wc_dir)
            lb.insert("end", f"  {rel}  :  {lineno}  |  {line[:80]}")

        def open_selected(event=None):
            sel = lb.curselection()
            if not sel:
                return
            fpath, lineno, _ = matches[sel[0]]
            # Find or load the doc
            existing = next((did for did, d in self.tree_state["docs"].items()
                             if d.get("path") == fpath), None)
            if existing:
                if existing not in self.tabs:
                    self.tabs.append(existing)
                self._switch_tab(existing)
            else:
                self._load_file(fpath)
            # Jump to the matching line
            self.root.after(80, lambda ln=lineno: (
                self.editor.mark_set("insert", f"{ln}.0"),
                self.editor.see(f"{ln}.0")))
            win.destroy()

        lb.bind("<Double-Button-1>", open_selected)
        tk.Button(win, text="Open Selected", bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=5, font=("Segoe UI", 10, "bold"),
                  command=open_selected).pack(side="right", padx=14, pady=(0,12))
        tk.Button(win, text="Close", bg=COLORS["bg3"], fg=COLORS["text1"],
                  relief="flat", padx=12, pady=5,
                  command=win.destroy).pack(side="right", pady=(0,12))

    # ── CLONE LINES ──────────────────────────────────────────────────────────
    def _clone_lines(self):
        try:
            sel_start = self.editor.index("sel.first")
            sel_end   = self.editor.index("sel.last")
        except tk.TclError:
            # No selection: use current line
            idx = self.editor.index("insert")
            ln = idx.split(".")[0]
            sel_start = f"{ln}.0"
            sel_end   = f"{ln}.end"

        # Expand to full lines
        ls_ln = sel_start.split(".")[0]
        le_ln = sel_end.split(".")[0]
        block_start = f"{ls_ln}.0"
        block_end   = f"{le_ln}.end"
        block = self.editor.get(block_start, block_end)

        # Insert after block_end
        self.editor.insert(block_end, "\n" + block)
        # Re-select the cloned block
        new_start = f"{int(le_ln)+1}.0"
        new_end   = f"{int(le_ln)+int(block.count(chr(10)))+1}.end"
        self.editor.tag_remove("sel", "1.0", "end")
        self.editor.mark_set("insert", new_end)
        self.editor.tag_add("sel", new_start, new_end)
        self.editor.see(new_end)

    # ── WRAP AS WILDCARD ─────────────────────────────────────────────────────
    def _wrap_wildcard(self):
        try:
            selected = self.editor.get("sel.first", "sel.last").strip()
        except tk.TclError:
            self._notify("Select text to wrap as wildcard", "warn")
            return
        if not selected: return
        wrap = self.cfg["wrap_str"]
        wrapped = wrap + selected + wrap
        # Snapshot current state before editing so Ctrl+Z restores exactly this
        self._flush_snap_timer()
        self._push_undo_snapshot()
        sel_start = self.editor.index("sel.first")
        sel_end   = self.editor.index("sel.last")
        self.editor.delete(sel_start, sel_end)
        self.editor.insert(sel_start, wrapped)

        # Check if wildcard exists
        exists = any(d["name"].lower() == selected.lower()
                     for d in self.tree_state["docs"].values())
        if not exists:
            # Also check on disk (including subfolders)
            disk_exists = False
            wc_dir = self.cfg.get("wc_dir","")
            if wc_dir and os.path.isdir(wc_dir):
                for root, dirs, files in os.walk(wc_dir):
                    for f in files:
                        if Path(f).stem.lower() == selected.lower():
                            disk_exists = True
                            break
                    if disk_exists: break
            if not disk_exists:
                # Create empty doc
                doc_id = str(uuid.uuid4())
                now = time.time()
                self.tree_state["docs"][doc_id] = {
                    "id": doc_id, "name": selected, "path": None,
                    "content": "", "color": None,
                    "created": now, "modified": now, "modified_unsaved": False
                }
                nwf = next((f for f in self.tree_state["folders"]
                            if f["name"].lower() == "new wildcards"), None)
                if not nwf:
                    nwf = {"id": str(uuid.uuid4()), "name": "new wildcards",
                           "color": COLORS["accent2"], "open": True,
                           "children": [], "docs": []}
                    self.tree_state["folders"].append(nwf)
                nwf["docs"].append(doc_id)
                self._notify(f'Created new wildcard: "{selected}"', "success")
                self._refresh_tree()
                save_tree_state(self.tree_state)

        self._apply_wildcard_highlights()
        self._update_wc_list()

    # ── NAV HISTORY ──────────────────────────────────────────────────────────
    def _nav_push(self, doc_id):
        self.nav_history = self.nav_history[:self.nav_index + 1]
        if self.nav_history and self.nav_history[-1] == doc_id:
            return
        self.nav_history.append(doc_id)
        self.nav_index = len(self.nav_history) - 1
        self._update_nav_buttons()

    def _nav_back(self):
        if self.nav_index <= 0: return
        self.nav_index -= 1
        did = self.nav_history[self.nav_index]
        if did in self.tree_state["docs"]:
            if did not in self.tabs: self.tabs.append(did)
            self._switch_tab(did, push_nav=False)
        self._update_nav_buttons()

    def _nav_forward(self):
        if self.nav_index >= len(self.nav_history) - 1: return
        self.nav_index += 1
        did = self.nav_history[self.nav_index]
        if did in self.tree_state["docs"]:
            if did not in self.tabs: self.tabs.append(did)
            self._switch_tab(did, push_nav=False)
        self._update_nav_buttons()

    def _update_nav_buttons(self):
        can_back = self.nav_index > 0
        can_fwd  = self.nav_index < len(self.nav_history) - 1
        self.btn_back.config(fg=COLORS["text1"] if can_back else COLORS["text2"],
                              cursor="hand2" if can_back else "arrow")
        self.btn_fwd.config( fg=COLORS["text1"] if can_fwd  else COLORS["text2"],
                              cursor="hand2" if can_fwd  else "arrow")

    # ── WORD WRAP ─────────────────────────────────────────────────────────────
    def _toggle_word_wrap(self):
        self.word_wrap.set(not self.word_wrap.get())
        mode = "word" if self.word_wrap.get() else "none"
        self.editor.config(wrap=mode)
        if self.word_wrap.get():
            self.ed_hscroll.pack_forget()
        else:
            self.ed_hscroll.pack(fill="x")
        self.cfg["word_wrap"] = self.word_wrap.get()
        save_config(self.cfg)
        self._update_wrap_btn()

    def _update_wrap_btn(self):
        state = "ON" if self.word_wrap.get() else "OFF"
        fg = COLORS["accent3"] if self.word_wrap.get() else COLORS["text2"]
        self.wrap_btn.config(text=f"⏎ Wrap: {state}", fg=fg)

    def _set_sidebar_sash(self):
        """Position the sidebar sash at 2/3 height for folder tree, 1/3 for wc list."""
        try:
            total = self.side_paned.winfo_height()
            if total > 30:
                self.side_paned.sash_place(0, 0, int(total * 2 / 3))
            else:
                # Not laid out yet, retry
                self.root.after(150, self._set_sidebar_sash)
        except Exception:
            pass

    # ── FOLDER TREE ───────────────────────────────────────────────────────────
    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self._tree_id_map = {}

        # Find ancestor folders of active doc for highlighting
        active_ancestors = set()
        if self.active_tab:
            self._find_ancestor_folders(self.active_tab, active_ancestors)

        def sort_key(item):
            mode = self.cfg.get("sort_mode", "name")
            if isinstance(item, dict):
                if mode == "name":     return item.get("name", "").lower()
                if mode == "created":  return item.get("created", 0)
                if mode == "modified": return item.get("modified", 0)
            return ""

        in_folder_docs = set()
        for f in self.tree_state["folders"]:
            for did in f.get("docs", []):
                in_folder_docs.add(did)

        child_folder_ids = set()
        for f in self.tree_state["folders"]:
            for cid in f.get("children", []):
                child_folder_ids.add(cid)

        top_folders = [f for f in self.tree_state["folders"]
                       if f["id"] not in child_folder_ids]
        top_folders.sort(key=sort_key)

        def _insert_doc(parent_iid, doc):
            in_folder_docs.add(doc["id"])
            is_active = (doc["id"] == self.active_tab)
            modified  = doc.get("modified_unsaved", False)
            doc_color = doc.get("color")
            tag_name  = f"d_{doc['id']}"

            base_label = "  📄 " + doc["name"] + (" *" if modified else "")
            # Line fill ONLY for the currently open file
            label = base_label + (" ─" * 40) if is_active else base_label

            if doc_color:
                # Color as text; subtle tint background; bold+accent if also active
                text_fg = COLORS["accent"] if is_active else doc_color
                tint = _color_tint(doc_color, 0.10)
                font_weight = "bold" if is_active else "normal"
                self.tree.tag_configure(tag_name, background=tint, foreground=text_fg,
                                         font=("Segoe UI", 10, font_weight))
            elif is_active:
                self.tree.tag_configure(tag_name, background=COLORS["sel_bg"],
                                         foreground=COLORS["accent"],
                                         font=("Segoe UI", 10, "bold"))
            else:
                self.tree.tag_configure(tag_name, background=COLORS["bg1"],
                                         foreground=COLORS["text1"], font=("Segoe UI", 10))
            diid = self.tree.insert(parent_iid, "end", text=label, tags=(tag_name,))
            self._tree_id_map[diid] = ("doc", doc["id"])

        def add_folder(parent_iid, folder, depth=0):
            is_open = folder.get("open", True)
            is_ancestor = folder["id"] in active_ancestors
            folder_color = folder.get("color")
            tag_name = f"f_{folder['id']}"
            # Line fill only on a collapsed ancestor folder (active doc hidden inside)
            is_collapsed_ancestor = is_ancestor and not is_open
            base_label = folder["name"]
            label = base_label + (" ─" * 40) if is_collapsed_ancestor else base_label

            if folder_color:
                text_fg = COLORS["accent"] if is_ancestor else folder_color
                tint = _color_tint(folder_color, 0.10)
                self.tree.tag_configure(tag_name, background=tint, foreground=text_fg,
                                         font=("Segoe UI", 10, "bold"))
            elif is_ancestor:
                self.tree.tag_configure(tag_name, background=COLORS["bg1"],
                                         foreground=COLORS["accent"], font=("Segoe UI", 10, "bold"))
            else:
                self.tree.tag_configure(tag_name, background=COLORS["bg1"],
                                         foreground=COLORS["text0"], font=("Segoe UI", 10, "bold"))
            iid = self.tree.insert(parent_iid, "end", text=label,
                                    open=is_open, tags=(tag_name,))
            self._tree_id_map[iid] = ("folder", folder["id"])
            children = [f for f in self.tree_state["folders"]
                        if f["id"] in folder.get("children", [])]
            children.sort(key=sort_key)
            for child in children:
                add_folder(iid, child, depth + 1)
            docs = [self.tree_state["docs"][did]
                    for did in folder.get("docs", [])
                    if did in self.tree_state["docs"]]
            docs.sort(key=sort_key)
            for doc in docs:
                _insert_doc(iid, doc)

        for folder in top_folders:
            add_folder("", folder)

        # Unsorted folder
        unsorted_open = self.tree_state.get("unsorted_open", True)
        unsorted_docs = [self.tree_state["docs"][did]
                         for did in self.tree_state["unsorted"]
                         if did in self.tree_state["docs"] and did not in in_folder_docs]
        unsorted_docs.sort(key=sort_key)
        is_u_ancestor = "_unsorted" in active_ancestors
        u_tag = "u_folder"
        if is_u_ancestor:
            self.tree.tag_configure(u_tag, background=COLORS["bg1"],
                                     foreground=COLORS["accent"], font=("Segoe UI", 10, "bold"))
        else:
            self.tree.tag_configure(u_tag, background=COLORS["bg1"],
                                     foreground=COLORS["text2"], font=("Segoe UI", 10, "bold"))
        u_iid = self.tree.insert("", "end", text="Unsorted", open=unsorted_open, tags=(u_tag,))
        self._tree_id_map[u_iid] = ("folder", "_unsorted")
        for doc in unsorted_docs:
            _insert_doc(u_iid, doc)

    def _find_ancestor_folders(self, doc_id, result_set):
        direct = set()
        for f in self.tree_state["folders"]:
            if doc_id in f.get("docs", []):
                direct.add(f["id"])
        if not direct:
            result_set.add("_unsorted")
            return
        result_set.update(direct)
        changed = True
        while changed:
            changed = False
            for f in self.tree_state["folders"]:
                for cid in f.get("children", []):
                    if cid in result_set and f["id"] not in result_set:
                        result_set.add(f["id"])
                        changed = True

    # ── TREE INTERACTION ─────────────────────────────────────────────────────
    def _tree_click(self, event):
        # If a drag just happened, process the drop instead
        if getattr(self, "_dragged", False):
            self._dragged = False
            self.tree.config(cursor="")
            self._drag_release(event)
            self._drag_data = None
            return
        iid = self.tree.identify_row(event.y)
        if not iid: return
        kind, obj_id = self._tree_id_map.get(iid, (None, None))
        if not kind: return

        now = time.time()

        if kind == "doc":
            # Slow double-click (within 0.55s on same item) = rename
            if iid == self._last_click_item and (now - self._last_click_time) < 0.55:
                self._rename_item(iid, kind, obj_id)
                self._last_click_item = None
                return
            self._last_click_item = iid
            self._last_click_time = now
            # Open doc on single click
            if obj_id not in self.tabs:
                self.tabs.append(obj_id)
            self._switch_tab(obj_id)

        elif kind == "folder" and obj_id != "_unsorted":
            # Slow double-click = rename
            if iid == self._last_click_item and (now - self._last_click_time) < 0.55:
                self._rename_item(iid, kind, obj_id)
                self._last_click_item = None
                return
            self._last_click_item = iid
            self._last_click_time = now

    def _tree_dbl_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid: return
        kind, obj_id = self._tree_id_map.get(iid, (None, None))
        # Docs are handled by single-click; double-click on doc = rename (handled in _tree_click)
        if kind == "folder":
            # Toggle open state and persist it
            cur = self.tree.item(iid, "open")
            new_open = not cur
            self.tree.item(iid, open=new_open)
            if obj_id == "_unsorted":
                self.tree_state["unsorted_open"] = new_open
            else:
                folder = next((f for f in self.tree_state["folders"]
                               if f["id"] == obj_id), None)
                if folder:
                    folder["open"] = new_open
            save_tree_state(self.tree_state)

    def _on_tree_open(self, event):
        iid = self.tree.focus()
        if not iid: return
        kind, obj_id = self._tree_id_map.get(iid, (None, None))
        if kind != "folder": return
        if obj_id == "_unsorted":
            self.tree_state["unsorted_open"] = True
        else:
            folder = next((f for f in self.tree_state["folders"] if f["id"] == obj_id), None)
            if folder: folder["open"] = True
        save_tree_state(self.tree_state)

    def _on_tree_close(self, event):
        iid = self.tree.focus()
        if not iid: return
        kind, obj_id = self._tree_id_map.get(iid, (None, None))
        if kind != "folder": return
        if obj_id == "_unsorted":
            self.tree_state["unsorted_open"] = False
        else:
            folder = next((f for f in self.tree_state["folders"] if f["id"] == obj_id), None)
            if folder: folder["open"] = False
        save_tree_state(self.tree_state)

    def _tree_right_click(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
        self.ctx_target = self._tree_id_map.get(iid, None)
        self._show_ctx_menu(event)

    def _rename_item(self, iid, kind, obj_id):
        if obj_id == "_unsorted": return
        if kind == "folder":
            obj = next((f for f in self.tree_state["folders"] if f["id"] == obj_id), None)
        else:
            obj = self.tree_state["docs"].get(obj_id)
        if not obj: return
        old_name = obj["name"]
        name = self._simple_input("Rename", "New name:", old_name)
        if name and name != old_name:
            obj["name"] = name
            if kind == "doc":
                self._rename_doc_file(obj, name)
                self._offer_wildcard_rename(old_name, name)
            self._refresh_tree()
            self._render_tabs()
            self.root.title(f"{obj['name']} — {APP_NAME}")
            save_tree_state(self.tree_state)

    # ── DRAG & DROP ──────────────────────────────────────────────────────────
    def _drag_start(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            self._drag_data = self._tree_id_map.get(iid)
            self._drag_source = iid
        else:
            self._drag_data = None

    def _drag_motion(self, event):
        self._dragged = True
        self.tree.config(cursor="fleur")

    def _drag_release(self, event):
        self.tree.config(cursor="")
        if not self._drag_data: return
        target_iid = self.tree.identify_row(event.y)
        if not target_iid or target_iid == getattr(self, "_drag_source", None):
            self._drag_data = None
            return
        src_kind, src_id = self._drag_data
        tgt_kind, tgt_id = self._tree_id_map.get(target_iid, (None, None))
        if not tgt_kind:
            self._drag_data = None
            return

        if src_kind == "doc":
            # Move doc to target folder
            if tgt_kind == "folder":
                self._move_doc_to_folder(src_id, tgt_id)
            elif tgt_kind == "doc":
                # Move to same folder as target
                for f in self.tree_state["folders"]:
                    if src_id in f.get("docs", []):
                        pass  # reorder (simplified: just move to same folder)
                # Find target's folder
                for f in self.tree_state["folders"]:
                    if tgt_id in f.get("docs", []):
                        self._move_doc_to_folder(src_id, f["id"])
                        break
        elif src_kind == "folder" and tgt_kind == "folder" and src_id != tgt_id:
            self._move_folder_into(src_id, tgt_id)

        self._drag_data = None
        self._refresh_tree()
        save_tree_state(self.tree_state)

    def _move_doc_to_folder(self, doc_id, folder_id):
        # Remove from all folders
        for f in self.tree_state["folders"]:
            if doc_id in f.get("docs", []):
                f["docs"].remove(doc_id)
        # Remove from unsorted
        if doc_id in self.tree_state["unsorted"]:
            self.tree_state["unsorted"].remove(doc_id)
        # Add to target
        if folder_id == "_unsorted":
            self.tree_state["unsorted"].append(doc_id)
        else:
            target = next((f for f in self.tree_state["folders"]
                           if f["id"] == folder_id), None)
            if target:
                target.setdefault("docs", []).append(doc_id)

    def _move_folder_into(self, src_id, dst_id):
        # Remove from parent children lists
        for f in self.tree_state["folders"]:
            if src_id in f.get("children", []):
                f["children"].remove(src_id)
        dst = next((f for f in self.tree_state["folders"] if f["id"] == dst_id), None)
        if dst and src_id not in dst.get("children", []):
            dst.setdefault("children", []).append(src_id)

    # ── CONTEXT MENU ─────────────────────────────────────────────────────────
    def _show_ctx_menu(self, event):
        menu = tk.Menu(self.root, tearoff=0,
                        bg=COLORS["bg2"], fg=COLORS["text0"],
                        activebackground=COLORS["bg3"],
                        activeforeground=COLORS["text0"],
                        relief="flat", bd=1)

        kind = self.ctx_target[0] if self.ctx_target else None
        obj_id = self.ctx_target[1] if self.ctx_target else None

        if kind == "doc":
            menu.add_command(label="Open", command=lambda: self._open_tab(obj_id))
        menu.add_command(label="Rename (slow dbl-click or F2)",
                         command=lambda: self._rename_via_ctx())
        menu.add_command(label="Set Color…",
                         command=lambda: self._color_dialog(kind, obj_id))
        menu.add_separator()
        menu.add_command(label="New Folder", command=self._new_folder_dialog)
        menu.add_separator()
        if kind == "doc":
            menu.add_command(label="Remove from List",
                             command=lambda: self._remove_from_list(obj_id),
                             foreground=COLORS["danger"])
        elif kind == "folder" and obj_id != "_unsorted":
            menu.add_command(label="Delete Folder",
                             command=lambda: self._delete_folder(obj_id),
                             foreground=COLORS["danger"])

        menu.post(event.x_root, event.y_root)

    def _rename_via_ctx(self):
        if not self.ctx_target: return
        kind, obj_id = self.ctx_target
        if kind == "folder":
            obj = next((f for f in self.tree_state["folders"] if f["id"] == obj_id), None)
        else:
            obj = self.tree_state["docs"].get(obj_id)
        if not obj: return
        old_name = obj["name"]
        name = self._simple_input("Rename", "New name:", old_name)
        if name and name != old_name:
            obj["name"] = name
            if kind == "doc":
                self._rename_doc_file(obj, name)
                self._offer_wildcard_rename(old_name, name)
            self._refresh_tree()
            self._render_tabs()
            save_tree_state(self.tree_state)

    def _rename_doc_file(self, doc, new_name):
        """Rename the .txt file on disk when a doc is renamed. Updates doc['path']."""
        old_path = doc.get("path")
        if not old_path or not os.path.isfile(old_path):
            return  # unsaved/in-memory doc — nothing to rename on disk
        old_dir  = os.path.dirname(old_path)
        new_path = os.path.join(old_dir, new_name + ".txt")
        if os.path.normpath(old_path) == os.path.normpath(new_path):
            return  # already correct
        try:
            os.rename(old_path, new_path)
            doc["path"] = new_path
        except Exception as e:
            messagebox.showerror("Rename Error",
                f"Could not rename file on disk:\n{old_path}\n→ {new_path}\n\n{e}",
                parent=self.root)

    def _offer_wildcard_rename(self, old_name, new_name):
        """After a doc rename, offer to update all wildcard references.
        Scans both the wc_dir on disk AND every in-memory open doc."""
        wrap    = self.cfg.get("wrap_str", "__")
        old_ref = wrap + old_name + wrap
        new_ref = wrap + new_name + wrap
        wc_dir  = self.cfg.get("wc_dir", "")

        # ── Collect affected items ────────────────────────────────────────────
        # Key: fpath (or None for unsaved in-memory docs), value: (content, doc_id_or_None)
        # We use an ordered dict keyed by canonical path to avoid duplicates.
        affected_by_path = {}   # fpath -> (content, doc_id)
        affected_unsaved = []   # (doc_id, content) for docs with no path

        # 1) Scan in-memory docs first — these are authoritative for open files
        for did, doc in self.tree_state["docs"].items():
            content = doc.get("content", "")
            # For the active tab, always read from the live editor widget
            if did == self.active_tab:
                content = self._get_real_content()
            if old_ref not in content:
                continue
            fpath = doc.get("path")
            if fpath:
                affected_by_path[os.path.normpath(fpath)] = (content, did)
            else:
                affected_unsaved.append((did, content))

        # 2) Scan disk for files NOT already represented in memory
        if wc_dir and os.path.isdir(wc_dir):
            try:
                for root_dir, dirs, files in os.walk(wc_dir):
                    for fname in sorted(files):
                        if not fname.endswith(".txt"):
                            continue
                        fpath = os.path.join(root_dir, fname)
                        npath = os.path.normpath(fpath)
                        if npath in affected_by_path:
                            continue  # already have in-memory version
                        # Also skip if this path belongs to a tracked doc (content already checked)
                        tracked = any(
                            os.path.normpath(d.get("path","")) == npath
                            for d in self.tree_state["docs"].values()
                        )
                        if tracked:
                            continue
                        try:
                            with open(fpath, "r", encoding="utf-8", errors="replace") as fh:
                                content = fh.read()
                            if old_ref in content:
                                affected_by_path[npath] = (content, None)
                        except Exception:
                            pass
            except Exception:
                pass

        total = len(affected_by_path) + len(affected_unsaved)
        if total == 0:
            return  # nothing to update, no popup needed

        ans = messagebox.askyesno(
            "Update Wildcard References",
            f"Found {total} file(s) referencing  {old_ref}\n\n"
            f"Replace all occurrences of:\n  {old_ref}\nwith:\n  {new_ref}\n\n"
            f"Save all affected files?",
            parent=self.root
        )
        if not ans:
            return

        updated = 0
        errors  = []

        # Update disk-backed files
        for npath, (content, did) in affected_by_path.items():
            new_content = content.replace(old_ref, new_ref)
            try:
                with open(npath, "w", encoding="utf-8") as fh:
                    fh.write(new_content)
                # Update in-memory doc if we have one
                if did and did in self.tree_state["docs"]:
                    self.tree_state["docs"][did]["content"] = new_content
                    self.tree_state["docs"][did]["saved_hash"] = self._content_hash(new_content)
                    if did == self.active_tab:
                        self.editor.delete("1.0", "end")
                        self.editor.insert("1.0", new_content)
                        self._apply_scroll_padding()
                        self.editor.edit_modified(False)
                else:
                    # Find any in-memory doc at this path and update it
                    for d_id, doc in self.tree_state["docs"].items():
                        if doc.get("path") and os.path.normpath(doc["path"]) == npath:
                            doc["content"] = new_content
                            doc["saved_hash"] = self._content_hash(new_content)
                            if d_id == self.active_tab:
                                self.editor.delete("1.0", "end")
                                self.editor.insert("1.0", new_content)
                                self._apply_scroll_padding()
                                self.editor.edit_modified(False)
                updated += 1
            except Exception as e:
                errors.append(f"{os.path.basename(npath)}: {e}")

        # Update unsaved in-memory-only docs
        for did, content in affected_unsaved:
            new_content = content.replace(old_ref, new_ref)
            if did in self.tree_state["docs"]:
                self.tree_state["docs"][did]["content"] = new_content
                if did == self.active_tab:
                    self.editor.delete("1.0", "end")
                    self.editor.insert("1.0", new_content)
                    self._apply_scroll_padding()
                    self.editor.edit_modified(False)
            updated += 1

        save_tree_state(self.tree_state)
        msg = f"Updated {updated} file(s)."
        if errors:
            msg += f"\n\nErrors ({len(errors)}):\n" + "\n".join(errors[:5])
        self._notify(msg, "success" if not errors else "warn")
        # Refresh highlights and wildcards-in-doc panel for the active tab
        self.root.after(50, self._apply_wildcard_highlights)
        self.root.after(50, self._update_wc_list)

    def _rename_current_doc(self):
        """Rename the currently active document (toolbar button handler)."""
        if not self.active_tab:
            return
        doc = self.tree_state["docs"].get(self.active_tab)
        if not doc:
            return
        old_name = doc["name"]
        name = self._simple_input("Rename", "New name:", old_name)
        if name and name != old_name:
            doc["name"] = name
            self._rename_doc_file(doc, name)
            self._offer_wildcard_rename(old_name, name)
            self._refresh_tree()
            self._render_tabs()
            self.root.title(f"{name} — {APP_NAME}")
            save_tree_state(self.tree_state)

    def _color_dialog(self, kind, obj_id):
        win = tk.Toplevel(self.root)
        win.title("Set Color")
        win.configure(bg=COLORS["bg2"])
        win.resizable(False, False)
        x = self.root.winfo_x() + 200
        y = self.root.winfo_y() + 200
        win.geometry(f"320x160+{x}+{y}")
        win.update_idletasks()
        win.grab_set()

        tk.Label(win, text="Choose a color:", bg=COLORS["bg2"],
                 fg=COLORS["text1"], font=("Segoe UI", 10)).pack(pady=(12,6))

        grid = tk.Frame(win, bg=COLORS["bg2"])
        grid.pack()
        sel = [None]
        btns = []
        for i, c in enumerate(PRESET_COLORS):
            b = tk.Label(grid, bg=c, width=2, height=1, cursor="hand2",
                         relief="flat", bd=2)
            b.grid(row=i//7, column=i%7, padx=2, pady=2)
            def on_click(color=c, btn=b):
                for x in btns: x.config(relief="flat")
                btn.config(relief="solid")
                sel[0] = color
            b.bind("<Button-1>", lambda e, fn=on_click: fn())
            btns.append(b)

        def apply():
            if sel[0]:
                if kind == "folder":
                    f = next((f for f in self.tree_state["folders"] if f["id"] == obj_id), None)
                    if f: f["color"] = sel[0]
                else:
                    doc = self.tree_state["docs"].get(obj_id)
                    if doc: doc["color"] = sel[0]
                self._refresh_tree()
                self._render_tabs()
                save_tree_state(self.tree_state)
            win.destroy()

        def clear_and_close():
            self._clear_color(kind, obj_id)
            win.destroy()

        btn_frame = tk.Frame(win, bg=COLORS["bg2"])
        btn_frame.pack(pady=8)
        tk.Button(btn_frame, text="Clear", bg=COLORS["bg3"], fg=COLORS["text2"],
                  relief="flat", command=clear_and_close).pack(side="left", padx=4)
        tk.Button(btn_frame, text="Apply", bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", font=("Segoe UI", 10, "bold"),
                  command=apply).pack(side="left", padx=4)

    def _clear_color(self, kind, obj_id):
        if kind == "folder":
            f = next((f for f in self.tree_state["folders"] if f["id"] == obj_id), None)
            if f: f["color"] = None
        else:
            doc = self.tree_state["docs"].get(obj_id)
            if doc: doc["color"] = None
        self._refresh_tree()
        self._render_tabs()
        save_tree_state(self.tree_state)

    def _remove_from_list(self, doc_id):
        for f in self.tree_state["folders"]:
            if doc_id in f.get("docs", []):
                f["docs"].remove(doc_id)
        if doc_id in self.tree_state["unsorted"]:
            self.tree_state["unsorted"].remove(doc_id)
        self._close_tab(doc_id)
        if doc_id in self.tree_state["docs"]:
            del self.tree_state["docs"][doc_id]
        self._refresh_tree()
        save_tree_state(self.tree_state)

    def _delete_folder(self, folder_id):
        folder = next((f for f in self.tree_state["folders"] if f["id"] == folder_id), None)
        if not folder: return

        has_contents = bool(folder.get("docs")) or bool(folder.get("children"))

        if has_contents:
            # Custom 3-button dialog
            win = tk.Toplevel(self.root)
            win.title("Delete Folder")
            win.configure(bg=COLORS["bg2"])
            win.resizable(False, False)
            x = self.root.winfo_x() + 160
            y = self.root.winfo_y() + 160
            win.geometry(f"380x170+{x}+{y}")
            win.update_idletasks()
            win.grab_set()

            tk.Label(win, text=f"Delete folder  \"{folder['name']}\"?",
                     bg=COLORS["bg2"], fg=COLORS["text0"],
                     font=("Segoe UI", 11, "bold")).pack(pady=(18, 4), padx=18, anchor="w")
            tk.Label(win, text="This folder has contents. What should happen to them?",
                     bg=COLORS["bg2"], fg=COLORS["text2"],
                     font=("Segoe UI", 9)).pack(padx=18, anchor="w")

            choice = tk.StringVar(value="cancel")

            btn_row = tk.Frame(win, bg=COLORS["bg2"])
            btn_row.pack(fill="x", padx=18, pady=20)

            def pick(val):
                choice.set(val)
                win.destroy()

            tk.Button(btn_row, text="🗑 Remove All",
                      bg=COLORS["danger"], fg="white",
                      relief="flat", padx=10, pady=6,
                      font=("Segoe UI", 9, "bold"),
                      command=lambda: pick("remove")).pack(side="left", padx=(0, 6))
            tk.Button(btn_row, text="⬇ Move to Unsorted",
                      bg=COLORS["bg4"], fg=COLORS["text0"],
                      relief="flat", padx=10, pady=6,
                      font=("Segoe UI", 9),
                      command=lambda: pick("unsorted")).pack(side="left", padx=(0, 6))
            tk.Button(btn_row, text="Cancel",
                      bg=COLORS["bg3"], fg=COLORS["text1"],
                      relief="flat", padx=10, pady=6,
                      font=("Segoe UI", 9),
                      command=lambda: pick("cancel")).pack(side="left")

            win.wait_window()
            action = choice.get()
            if action == "cancel":
                return
        else:
            action = "remove"  # empty folder — just delete it silently

        def _collect_all_docs_in_folder(fid):
            """Recursively collect all doc IDs within a folder and its children."""
            f = next((x for x in self.tree_state["folders"] if x["id"] == fid), None)
            if not f:
                return []
            docs = list(f.get("docs", []))
            for cid in f.get("children", []):
                docs.extend(_collect_all_docs_in_folder(cid))
            return docs

        def _remove_folder_recursive(fid):
            """Remove a folder and all its child folders from tree_state."""
            f = next((x for x in self.tree_state["folders"] if x["id"] == fid), None)
            if not f:
                return
            for cid in list(f.get("children", [])):
                _remove_folder_recursive(cid)
            self.tree_state["folders"] = [x for x in self.tree_state["folders"] if x["id"] != fid]

        all_docs = _collect_all_docs_in_folder(folder_id)

        if action == "remove":
            # Close open tabs and delete docs entirely
            for did in all_docs:
                if did in self.tabs:
                    self.tabs.remove(did)
                    if self.active_tab == did:
                        self.active_tab = self.tabs[0] if self.tabs else None
                if did in self.tree_state["docs"]:
                    del self.tree_state["docs"][did]
                if did in self.tree_state["unsorted"]:
                    self.tree_state["unsorted"].remove(did)
            if self.active_tab:
                self._switch_tab(self.active_tab, push_nav=False)
            else:
                self.editor.delete("1.0", "end")
        elif action == "unsorted":
            # Move all docs to unsorted
            for did in all_docs:
                if did not in self.tree_state["unsorted"]:
                    self.tree_state["unsorted"].append(did)

        # Remove the folder and all its descendants from the folders list
        _remove_folder_recursive(folder_id)
        # Unlink from any parent's children list
        for f in self.tree_state["folders"]:
            f["children"] = [c for c in f.get("children", []) if c != folder_id]

        self._render_tabs()
        self._refresh_tree()
        save_tree_state(self.tree_state)

    # ── NEW FOLDER ────────────────────────────────────────────────────────────
    def _import_folder_structure(self):
        """Let the user pick a directory, then walk it and import all folders and
        .txt files into the explorer, preserving the full subfolder hierarchy.
        Files already tracked (same path) are skipped to avoid duplicates."""

        import_dir = filedialog.askdirectory(
            title="Select folder to import",
            initialdir=self.cfg.get("wc_dir") or os.path.expanduser("~"))
        if not import_dir or not os.path.isdir(import_dir):
            return

        # Build a fast lookup of already-tracked paths so we never double-import
        existing_paths = {doc.get("path"): did
                          for did, doc in self.tree_state["docs"].items()
                          if doc.get("path")}

        # Also build a lookup of folder objects by their disk path so we can
        # find/create folders without duplicating them
        # We'll populate this as we walk the tree

        added_folders = 0
        added_docs    = 0

        def get_or_create_folder(dir_path, parent_folder_id=None):
            """Return the folder-state dict for dir_path, creating it if needed.
            parent_folder_id=None means it's a top-level folder."""
            nonlocal added_folders
            folder_name = os.path.basename(dir_path)

            # Check whether this folder already exists in the tree by matching
            # name + parent relationship (avoids duplicating same-named folders)
            if parent_folder_id is None:
                # Top-level: its id must NOT appear in any children list
                child_ids = {cid for f in self.tree_state["folders"]
                             for cid in f.get("children", [])}
                candidates = [f for f in self.tree_state["folders"]
                              if f["name"] == folder_name
                              and f["id"] not in child_ids]
            else:
                parent = next((f for f in self.tree_state["folders"]
                               if f["id"] == parent_folder_id), None)
                child_ids_of_parent = set(parent.get("children", [])) if parent else set()
                candidates = [f for f in self.tree_state["folders"]
                              if f["id"] in child_ids_of_parent
                              and f["name"] == folder_name]

            if candidates:
                return candidates[0]

            # Create a new folder entry
            fid = str(uuid.uuid4())
            new_folder = {
                "id": fid, "name": folder_name, "color": None,
                "open": True, "children": [], "docs": []
            }
            self.tree_state["folders"].append(new_folder)
            added_folders += 1

            # Link to parent
            if parent_folder_id is not None:
                parent = next((f for f in self.tree_state["folders"]
                               if f["id"] == parent_folder_id), None)
                if parent and fid not in parent.get("children", []):
                    parent.setdefault("children", []).append(fid)

            return new_folder

        def walk_dir(dir_path, parent_folder_id=None, is_root=False):
            nonlocal added_docs

            try:
                entries = sorted(os.scandir(dir_path), key=lambda e: (not e.is_dir(), e.name.lower()))
            except PermissionError:
                return

            # Determine this directory's folder object (None if root — we don't create one)
            if is_root:
                this_folder_obj = None
                this_folder_id  = None
            else:
                this_folder_obj = get_or_create_folder(dir_path, parent_folder_id)
                this_folder_id  = this_folder_obj["id"]

            for entry in entries:
                if entry.is_dir(follow_symlinks=False):
                    # Subdirs of root become top-level folders (parent_folder_id=None)
                    # Subdirs of a proper folder nest inside it
                    walk_dir(entry.path, parent_folder_id=this_folder_id, is_root=False)
                elif entry.is_file() and entry.name.lower().endswith(".txt"):
                    dest_folder_obj = this_folder_obj  # None for root-level files → Unsorted
                    file_name = Path(entry.path).stem

                    # Skip if already tracked by exact path
                    if entry.path in existing_paths:
                        did = existing_paths[entry.path]
                        if dest_folder_obj is None:
                            for f in self.tree_state["folders"]:
                                if did in f.get("docs", []):
                                    f["docs"].remove(did)
                            if did not in self.tree_state["unsorted"]:
                                self.tree_state["unsorted"].append(did)
                        else:
                            if did not in dest_folder_obj.get("docs", []):
                                for f in self.tree_state["folders"]:
                                    if did in f.get("docs", []) and f["id"] != dest_folder_obj["id"]:
                                        f["docs"].remove(did)
                                if did in self.tree_state["unsorted"]:
                                    self.tree_state["unsorted"].remove(did)
                                if did not in dest_folder_obj["docs"]:
                                    dest_folder_obj["docs"].append(did)
                        continue

                    # Skip if a doc with the same name already exists in the target location
                    if dest_folder_obj is None:
                        # Check unsorted for same name
                        name_exists = any(
                            self.tree_state["docs"].get(did, {}).get("name", "").lower() == file_name.lower()
                            for did in self.tree_state["unsorted"]
                        )
                    else:
                        # Check destination folder for same name
                        name_exists = any(
                            self.tree_state["docs"].get(did, {}).get("name", "").lower() == file_name.lower()
                            for did in dest_folder_obj.get("docs", [])
                        )
                    if name_exists:
                        continue

                    # Load the file
                    try:
                        with open(entry.path, "r", encoding="utf-8", errors="replace") as fh:
                            content = fh.read()
                    except Exception:
                        continue

                    st     = os.stat(entry.path)
                    doc_id = str(uuid.uuid4())
                    name   = Path(entry.path).stem
                    self.tree_state["docs"][doc_id] = {
                        "id": doc_id, "name": name, "path": entry.path,
                        "content": content, "color": None,
                        "created": st.st_ctime, "modified": st.st_mtime,
                        "modified_unsaved": False,
                        "saved_hash": self._content_hash(content),
                    }
                    if dest_folder_obj is None:
                        self.tree_state["unsorted"].append(doc_id)
                    else:
                        dest_folder_obj["docs"].append(doc_id)
                    existing_paths[entry.path] = doc_id
                    added_docs += 1

        # Walk the chosen directory: its immediate files → Unsorted,
        # its subdirectories → top-level explorer folders (recursively)
        walk_dir(import_dir, parent_folder_id=None, is_root=True)

        save_tree_state(self.tree_state)
        self._refresh_tree()
        self._render_tabs()

        msg = f"Imported {added_docs} file(s) into {added_folders} new folder(s)."
        self._notify(msg, "success")
        if added_docs == 0 and added_folders == 0:
            messagebox.showinfo("Import Complete",
                "Nothing new was imported — all files in that folder are already in the explorer.",
                parent=self.root)
        else:
            messagebox.showinfo("Import Complete", msg, parent=self.root)

    def _new_folder_dialog(self):
        name = self._simple_input("New Folder", "Folder name:", "")
        if not name: return
        fid = str(uuid.uuid4())
        self.tree_state["folders"].append({
            "id": fid, "name": name, "color": None,
            "open": True, "children": [], "docs": []
        })
        self._refresh_tree()
        save_tree_state(self.tree_state)
        self._notify(f"Created folder: {name}", "success")

    # ── SORT ─────────────────────────────────────────────────────────────────
    def _show_sort_menu(self):
        menu = tk.Menu(self.root, tearoff=0,
                        bg=COLORS["bg2"], fg=COLORS["text0"],
                        activebackground=COLORS["bg3"],
                        activeforeground=COLORS["text0"],
                        relief="flat")
        cur = self.cfg.get("sort_mode","name")
        for mode, label in [("name","By Name"),("created","By Created Date"),("modified","By Modified Date")]:
            pfx = "✓ " if cur==mode else "  "
            menu.add_command(label=pfx+label,
                             command=lambda m=mode: self._set_sort(m))
        # Position near top of sidebar
        x = self.sidebar.winfo_rootx() + 60
        y = self.sidebar.winfo_rooty() + 30
        menu.post(x, y)

    def _set_sort(self, mode):
        self.cfg["sort_mode"] = mode
        save_config(self.cfg)
        self._refresh_tree()

    # ── FIND & REPLACE ────────────────────────────────────────────────────────
    def _toggle_find(self):
        if self.find_frame.winfo_ismapped():
            self.find_frame.pack_forget()
        else:
            self.find_frame.pack(fill="x", before=self.editor_wrap)
            self.find_entry.focus_set()

    def _open_find_with_selection(self):
        """Open find panel, and if text is selected paste it into the find field."""
        try:
            sel = self.editor.get("sel.first", "sel.last")
        except tk.TclError:
            sel = ""
        if not self.find_frame.winfo_ismapped():
            self.find_frame.pack(fill="x", before=self.editor_wrap)
        if sel:
            self.find_var.set(sel)
            self.find_entry.select_range(0, "end")
        self.find_entry.focus_set()
        self._do_find_highlight()

    def _set_find_mode(self, mode):
        self.find_mode = mode
        for m, btn in self._mode_btns.items():
            btn.config(bg=COLORS["accent"] if m==mode else COLORS["bg3"],
                       fg=COLORS["bg0"] if m==mode else COLORS["text2"])
        self._do_find_highlight()

    def _build_find_regex(self):
        q = self.find_var.get()
        if not q: return None
        if self.find_mode == "extended":
            q = q.replace("\\n","\n").replace("\\t","\t").replace("\\r","\r")
        flags = 0 if self.find_case.get() else re.IGNORECASE
        if self.find_mode != "regex":
            q = re.escape(q)
        if self.find_whole.get():
            q = r"\b" + q + r"\b"
        try:
            return re.compile(q, flags)
        except re.error:
            return None

    def _do_find_highlight(self, *args, keep_pos=None):
        if not hasattr(self, "editor"):
            return
        self.editor.tag_remove("find_hl", "1.0", "end")
        self.editor.tag_remove("find_cur", "1.0", "end")
        self.find_matches = []
        self.find_current = -1
        rx = self._build_find_regex()
        if not rx:
            self.find_status.config(text="")
            return
        content = self._get_real_content()
        for m in rx.finditer(content):
            s = f"1.0+{m.start()}c"
            e = f"1.0+{m.end()}c"
            self.find_matches.append((s, e))
            self.editor.tag_add("find_hl", s, e)
        n = len(self.find_matches)
        if n == 0:
            self.find_status.config(text="No results", fg=COLORS["danger"])
            self.find_entry.config(highlightcolor=COLORS["danger"],
                                   highlightbackground=COLORS["danger"])
        else:
            self.find_status.config(text=f"0/{n}", fg=COLORS["text2"])
            self.find_entry.config(highlightcolor=COLORS["accent"],
                                   highlightbackground=COLORS["border2"])
            if keep_pos is not None:
                # Advance to the match at or after keep_pos (used by replace_current)
                self.find_current = min(keep_pos, n - 1)
                self._select_match()
            elif self.find_current < 0:
                self._find_next()

    def _find_next(self):
        if not self.find_matches: self._do_find_highlight(); return
        self.find_current = (self.find_current + 1) % len(self.find_matches)
        self._select_match()

    def _find_prev(self):
        if not self.find_matches: self._do_find_highlight(); return
        self.find_current = (self.find_current - 1) % len(self.find_matches)
        self._select_match()

    def _select_match(self):
        self.editor.tag_remove("find_cur", "1.0", "end")
        if self.find_current < 0 or self.find_current >= len(self.find_matches):
            return
        s, e = self.find_matches[self.find_current]
        self.editor.tag_add("find_cur", s, e)
        self.editor.see(s)
        self.editor.mark_set("insert", s)
        n = len(self.find_matches)
        self.find_status.config(text=f"{self.find_current+1}/{n}", fg=COLORS["text2"])

    def _replace_current(self):
        if self.find_current < 0 or not self.find_matches:
            return
        # Remember scroll position and cursor before doing anything
        try:
            view_pos = self.editor.yview()[0]
            cursor_before = self.editor.index("insert")
        except Exception:
            view_pos = 0.0
            cursor_before = "1.0"

        # Snapshot before the change so it's undoable
        self._push_undo_snapshot()
        s, e = self.find_matches[self.find_current]
        r = self.replace_var.get()
        prev_idx = self.find_current

        self._undo_inhibit = True
        try:
            self.editor.delete(s, e)
            self.editor.insert(s, r)
            did = self.active_tab
            if did and did in self.tree_state["docs"]:
                self.tree_state["docs"][did]["content"] = self._get_real_content()
                self.tree_state["docs"][did]["modified_unsaved"] = True
        finally:
            self._undo_inhibit = False

        self._push_undo_snapshot()

        # Rebuild match list but DON'T move the cursor or scroll — stay in place
        self.editor.tag_remove("find_hl", "1.0", "end")
        self.editor.tag_remove("find_cur", "1.0", "end")
        self.find_matches = []
        self.find_current = -1
        rx = self._build_find_regex()
        if rx:
            content = self._get_real_content()
            for m in rx.finditer(content):
                ms = f"1.0+{m.start()}c"
                me = f"1.0+{m.end()}c"
                self.find_matches.append((ms, me))
                self.editor.tag_add("find_hl", ms, me)
        n = len(self.find_matches)
        if n == 0:
            self.find_status.config(text="No results", fg=COLORS["danger"])
            self.find_entry.config(highlightcolor=COLORS["danger"],
                                   highlightbackground=COLORS["danger"])
        else:
            # Highlight the next match (or last if we replaced the final one)
            # but do NOT scroll or move the cursor
            self.find_current = min(prev_idx, n - 1)
            ms, me = self.find_matches[self.find_current]
            self.editor.tag_add("find_cur", ms, me)
            self.find_status.config(text=f"{self.find_current+1}/{n}", fg=COLORS["text2"])
            self.find_entry.config(highlightcolor=COLORS["accent"],
                                   highlightbackground=COLORS["border2"])

        # Restore scroll position and cursor
        try:
            self.editor.yview_moveto(view_pos)
            self.editor.mark_set("insert", cursor_before)
        except Exception:
            pass

    def _replace_all(self):
        rx = self._build_find_regex()
        if not rx:
            return
        r = self.replace_var.get()
        content = self._get_real_content()
        new_content, count = rx.subn(r, content)
        if count:
            # Snapshot before so the entire replace-all is one undo step
            self._push_undo_snapshot()
            self._undo_inhibit = True
            try:
                self._remove_scroll_padding()
                self.editor.config(undo=False)
                self.editor.delete("1.0", "end")
                self.editor.insert("1.0", new_content)
                self.editor.config(undo=True)
                self.editor.edit_modified(False)
                self._apply_scroll_padding()
                did = self.active_tab
                if did and did in self.tree_state["docs"]:
                    self.tree_state["docs"][did]["content"] = new_content
                    self.tree_state["docs"][did]["modified_unsaved"] = True
            finally:
                self._undo_inhibit = False
            # Snapshot after so redo can return to this state
            self._push_undo_snapshot()
            self._notify(f"Replaced {count} occurrence{'s' if count!=1 else ''}", "success")
        self._do_find_highlight()

    # ── SPELL CHECK ───────────────────────────────────────────────────────────
    def _get_spell(self):
        """Return a cached SpellChecker instance, or None if unavailable.
        Never caches a None result — retries on every call if not yet found."""
        # Return cached working instance
        if getattr(self, "_spell_instance", None) is not None:
            return self._spell_instance

        spell = None

        # Strategy 1: direct import (works if pip install used the same Python)
        try:
            from spellchecker import SpellChecker
            spell = SpellChecker()
        except Exception:
            pass

        # Strategy 2: if direct import failed, add site-packages from the
        # actual running interpreter and retry
        if spell is None:
            import sys, subprocess
            try:
                result = subprocess.run(
                    [sys.executable, "-c",
                     "import spellchecker; print(spellchecker.__file__)"],
                    capture_output=True, text=True, timeout=5
                )
                if result.returncode == 0 and result.stdout.strip():
                    pkg_path = str(Path(result.stdout.strip()).parent.parent)
                    if pkg_path not in sys.path:
                        sys.path.insert(0, pkg_path)
                    # Retry import with updated path
                    import importlib
                    sc_mod = importlib.import_module("spellchecker")
                    spell = sc_mod.SpellChecker()
            except Exception:
                pass

        if spell is None:
            return None  # Don't cache — allow retry next time

        # Load user dictionary
        dict_path = CONFIG_PATH.parent / "user_dict.txt"
        if dict_path.exists():
            try:
                spell.word_frequency.load_text_file(str(dict_path))
            except Exception:
                pass

        self._spell_instance = spell
        return spell

    def _toggle_spell(self):
        self.spell_enabled = not self.spell_enabled
        self.cfg["spell_check"] = self.spell_enabled
        save_config(self.cfg)
        if self.spell_enabled:
            # Let _run_spell_check set the label — it sets ✓ on success, error on failure
            self._run_spell_check()
        else:
            self.sb_spell.config(text="Spell ✗", fg=COLORS["text2"])
            self.editor.tag_remove("spell_err", "1.0", "end")

    def _run_spell_check(self):
        if not self.spell_enabled: return
        spell = self._get_spell()
        if spell is None:
            self.sb_spell.config(text="Spell: pip install pyspellchecker", fg=COLORS["warn"])
            return
        self.sb_spell.config(text="Spell ✓", fg=COLORS["accent3"])
        self.editor.tag_remove("spell_err", "1.0", "end")
        content = self._get_real_content()
        word_positions = [(m.group(0), m.start(), m.end())
                          for m in re.finditer(r"\b[a-zA-Z]{2,}\b", content)]
        if not word_positions: return
        unique = {w.lower() for w, _, _ in word_positions}
        try:
            misspelled = spell.unknown(unique)
        except Exception:
            return
        wrap = re.escape(self.cfg.get("wrap_str", "__"))
        wc_names = {m.group(1).lower()
                    for m in re.finditer(wrap + r"([^\s]+?)" + wrap, content)}
        for word, start, end in word_positions:
            if word.lower() in misspelled and word.lower() not in wc_names:
                self.editor.tag_add("spell_err", f"1.0+{start}c", f"1.0+{end}c")

    def _editor_right_click(self, event):
        """Unified right-click menu: general editing actions + spell suggestions if applicable."""
        menu = tk.Menu(self.root, tearoff=0,
                       bg=COLORS["bg2"], fg=COLORS["text0"],
                       activebackground=COLORS["accent"],
                       activeforeground=COLORS["bg0"],
                       font=("Segoe UI", 10))

        # ── Selection state ───────────────────────────────────────────────────
        try:
            sel_start = self.editor.index("sel.first")
            sel_end   = self.editor.index("sel.last")
            has_sel   = True
        except tk.TclError:
            has_sel = False

        # ── Clipboard state ───────────────────────────────────────────────────
        try:
            clip = self.root.clipboard_get()
            has_clip = bool(clip)
        except Exception:
            has_clip = False

        def cmd(label, fn, enabled=True):
            state = "normal" if enabled else "disabled"
            menu.add_command(label=label, command=fn, state=state)

        # Cut / Copy / Paste / Delete
        cmd("Cut",        lambda: self.editor.event_generate("<<Cut>>"),   has_sel)
        cmd("Copy",       lambda: self.editor.event_generate("<<Copy>>"),  has_sel)
        cmd("Paste",      lambda: self._paste_real(), has_clip)
        cmd("Delete",     lambda: self.editor.delete("sel.first","sel.last"), has_sel)
        menu.add_separator()

        # Select All
        cmd("Select All", lambda: (self.editor.tag_add("sel","1.0","end-1c"),
                                   self.editor.mark_set("insert","end-1c")))
        menu.add_separator()

        # Undo / Redo
        cmd("Undo", self._do_undo)
        cmd("Redo", self._do_redo)
        menu.add_separator()

        # Wildcard actions
        cmd("Wrap as Wildcard",  self._wrap_wildcard, has_sel)
        cmd("Clone Line(s)",     self._clone_lines)
        menu.add_separator()

        # Find
        cmd("Find…",             lambda: self._open_find_with_selection())
        menu.add_separator()

        # ── Spell check section (only if on a misspelled word) ────────────────
        idx = self.editor.index(f"@{event.x},{event.y}")
        if self.spell_enabled and "spell_err" in self.editor.tag_names(idx):
            word_start = self.editor.index(f"{idx} wordstart")
            word_end   = self.editor.index(f"{idx} wordend")
            word       = self.editor.get(word_start, word_end).strip()
            if word and word.isalpha():
                spell = self._get_spell()
                candidates = []
                if spell:
                    try:
                        raw = spell.candidates(word.lower())
                        candidates = sorted(raw or [])[:8]
                    except Exception:
                        pass

                menu.add_command(label=f'Misspelled: "{word}"',
                                 state="disabled", foreground=COLORS["danger"])
                menu.add_separator()
                if candidates:
                    for suggestion in candidates:
                        def _replace(s=suggestion, ws=word_start, we=word_end, orig=word):
                            if orig[0].isupper():
                                s = s[0].upper() + s[1:]
                            self.editor.delete(ws, we)
                            self.editor.insert(ws, s)
                            self.root.after(200, self._run_spell_check)
                        menu.add_command(label=f"  → {suggestion}", command=_replace)
                else:
                    menu.add_command(label="  (no suggestions)", state="disabled")
                menu.add_separator()
                menu.add_command(label="Add to dictionary",
                                 command=lambda w=word.lower(): self._add_to_dict(w))
                menu.add_separator()

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _paste_real(self):
        """Paste clipboard, replacing selection if present."""
        try:
            self.editor.delete("sel.first", "sel.last")
        except tk.TclError:
            pass
        try:
            self.editor.insert("insert", self.root.clipboard_get())
        except Exception:
            pass

    def _spell_right_click(self, event):
        """Legacy alias — now handled by _editor_right_click."""
        self._editor_right_click(event)


        """Right-click on a misspelled word: show correction menu."""
        if not self.spell_enabled: return
        idx = self.editor.index(f"@{event.x},{event.y}")
        if "spell_err" not in self.editor.tag_names(idx): return
        word_start = self.editor.index(f"{idx} wordstart")
        word_end   = self.editor.index(f"{idx} wordend")
        word = self.editor.get(word_start, word_end).strip()
        if not word or not word.isalpha(): return
        spell = self._get_spell()
        # Get suggestions — candidates() can return None for some words
        candidates = []
        if spell:
            try:
                raw = spell.candidates(word.lower())
                candidates = sorted(raw or [])[:8]
            except Exception:
                candidates = []
        menu = tk.Menu(self.root, tearoff=0,
                        bg=COLORS["bg2"], fg=COLORS["text0"],
                        activebackground=COLORS["accent"],
                        activeforeground=COLORS["bg0"],
                        font=("Segoe UI", 10))
        menu.add_command(label=f'Misspelled: "{word}"', state="disabled",
                         foreground=COLORS["danger"])
        menu.add_separator()
        if candidates:
            for suggestion in candidates:
                def replace(s=suggestion, ws=word_start, we=word_end, orig=word):
                    if orig[0].isupper():
                        s = s[0].upper() + s[1:]
                    self.editor.delete(ws, we)
                    self.editor.insert(ws, s)
                    self.root.after(200, self._run_spell_check)
                menu.add_command(label=suggestion, command=replace)
        else:
            menu.add_command(label="(no suggestions)", state="disabled")
        menu.add_separator()
        menu.add_command(label="Add to dictionary",
                         command=lambda w=word.lower(): self._add_to_dict(w))
        try:
            menu.post(event.x_root, event.y_root)
        except Exception:
            pass

    def _add_to_dict(self, word):
        dict_path = CONFIG_PATH.parent / "user_dict.txt"
        try:
            with open(dict_path, "a", encoding="utf-8") as f:
                f.write(word + "\n")
            # Reset cache so dict reloads
            if hasattr(self, "_spell_instance"):
                del self._spell_instance
        except Exception:
            pass
        self._run_spell_check()

    # ── REORGANIZE ────────────────────────────────────────────────────────────
    # ── WRAPPER INTEGRITY CHECK ───────────────────────────────────────────────
    def _check_wrapper_integrity(self, content):
        """Scan content for malformed wildcard wrapper usage.
        Returns True if problems found (and user chose not to save), False if clean/ignored."""
        wrap = self.cfg.get("wrap_str", "__")
        if not wrap:
            return False
        ch = wrap[0]  # e.g. '~'
        # Build pattern: any run of the wrapper char that isn't exactly len(wrap)
        # e.g. for ~~ we flag: single ~ not part of ~~, or ~~~+
        w = len(wrap)
        # Match runs of ch of length != w (i.e. 1..w-1 or w+1..)
        pattern = re.compile(re.escape(ch) + "+")
        problems = []
        for m in re.finditer(pattern, content):
            run = m.group()
            if len(run) % w != 0:
                problems.append((m.start(), m.end(), run))

        if not problems:
            self.editor.tag_remove("warn_tilde", "1.0", "end")
            return False

        # Highlight all problem spans
        self.editor.tag_remove("warn_tilde", "1.0", "end")
        first_pos = None
        for start, end, run in problems:
            s = f"1.0+{start}c"
            e = f"1.0+{end}c"
            self.editor.tag_add("warn_tilde", s, e)
            if first_pos is None:
                first_pos = s
        # Raise warn_tilde above other tags
        try:
            self.editor.tag_raise("warn_tilde")
            self.editor.tag_raise("sel")
        except Exception:
            pass
        # Scroll to first problem
        if first_pos:
            self.editor.see(first_pos)

        desc_lines = []
        for start, end, run in problems[:5]:
            # Find line number
            line = content[:start].count("\n") + 1
            desc_lines.append(f"  Line {line}: {''.join(run)!r}")
        if len(problems) > 5:
            desc_lines.append(f"  … and {len(problems)-5} more")

        ans = messagebox.askyesno(
            "Wildcard Wrapper Warning",
            f"Found {len(problems)} malformed wildcard wrapper(s) using '{wrap}':\n\n"
            + "\n".join(desc_lines)
            + f"\n\nA stray '{ch}' or '{ch*3}' can break wildcard parsing.\n\n"
            "Save anyway?",
            parent=self.root
        )
        return not ans  # True = abort save, False = proceed

    # ── REMOVE ISOLATED WILDCARDS ─────────────────────────────────────────────
    def _remove_isolated_wildcards(self):
        """Remove from the explorer any wildcard doc that is neither called by
        nor calls any other wildcard doc present in the wildcards folder."""
        wc_dir = self.cfg.get("wc_dir", "")
        wrap   = self.cfg.get("wrap_str", "__")
        esc    = re.escape(wrap)
        ref_pattern = re.compile(esc + r"([^\s]+?)" + esc)

        # Build set of all doc names (lowercased) in the explorer
        all_names = {doc["name"].lower(): did
                     for did, doc in self.tree_state["docs"].items()}

        # For each doc, collect the set of wildcard names it references
        def get_refs(doc):
            content = doc.get("content", "")
            return {m.group(1).lower() for m in ref_pattern.finditer(content)}

        # Build call graph
        calls   = {}  # doc_id -> set of names this doc references
        called_by = {did: set() for did in self.tree_state["docs"]}  # doc_id -> set of doc_ids that call it

        for did, doc in self.tree_state["docs"].items():
            refs = get_refs(doc)
            calls[did] = refs
            for ref_name in refs:
                target_id = all_names.get(ref_name)
                if target_id:
                    called_by[target_id].add(did)

        # Isolated = not called by anyone AND calls no one present in the explorer
        isolated = []
        for did, doc in self.tree_state["docs"].items():
            no_callers = len(called_by.get(did, set())) == 0
            no_callees = not any(name in all_names for name in calls.get(did, set()))
            if no_callers and no_callees:
                isolated.append(did)

        if not isolated:
            messagebox.showinfo("Remove Isolated", "No isolated wildcards found.", parent=self.root)
            return

        # Show confirmation with list
        names_preview = ", ".join(
            self.tree_state["docs"][did]["name"] for did in isolated[:10])
        if len(isolated) > 10:
            names_preview += f" … and {len(isolated)-10} more"

        win = tk.Toplevel(self.root)
        win.title("Remove Isolated Wildcards")
        win.configure(bg=COLORS["bg2"])
        win.resizable(True, True)
        x = self.root.winfo_x() + 80
        y = self.root.winfo_y() + 80
        win.geometry(f"520x400+{x}+{y}")
        win.update_idletasks()
        win.grab_set()

        tk.Label(win, text=f"Found {len(isolated)} isolated wildcard(s)",
                 bg=COLORS["bg2"], fg=COLORS["warn"],
                 font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=14, pady=(12,2))
        tk.Label(win, text="These docs neither call nor are called by any other wildcard in the explorer.",
                 bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9), wraplength=490).pack(anchor="w", padx=14)

        list_frame = tk.Frame(win, bg=COLORS["bg1"])
        list_frame.pack(fill="both", expand=True, padx=14, pady=8)
        sb = ttk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")
        lb = tk.Listbox(list_frame, bg=COLORS["bg1"], fg=COLORS["text1"],
                         selectbackground=COLORS["sel_bg"],
                         font=("Segoe UI", 10), relief="flat", bd=0,
                         activestyle="none", selectmode="extended",
                         yscrollcommand=sb.set)
        lb.pack(fill="both", expand=True)
        sb.config(command=lb.yview)
        for did in isolated:
            lb.insert("end", "  " + self.tree_state["docs"][did]["name"])
        lb.select_set(0, "end")

        choice = tk.StringVar(value="cancel")
        def pick(v): choice.set(v); win.destroy()

        btn_row = tk.Frame(win, bg=COLORS["bg2"])
        btn_row.pack(fill="x", padx=14, pady=(0, 12))
        tk.Button(btn_row, text="Remove Selected from Explorer",
                  bg=COLORS["danger"], fg="white", relief="flat",
                  padx=10, pady=5, font=("Segoe UI", 9, "bold"),
                  command=lambda: pick("remove")).pack(side="left", padx=(0,6))
        tk.Button(btn_row, text="Cancel",
                  bg=COLORS["bg3"], fg=COLORS["text1"], relief="flat",
                  padx=10, pady=5, command=lambda: pick("cancel")).pack(side="left")

        win.wait_window()
        if choice.get() != "remove":
            return

        selected_indices = lb.curselection() if lb.winfo_exists() else range(len(isolated))
        to_remove = [isolated[i] for i in range(len(isolated))]

        removed = 0
        for did in to_remove:
            # Remove from folders
            for f in self.tree_state["folders"]:
                if did in f.get("docs", []):
                    f["docs"].remove(did)
            # Remove from unsorted
            if did in self.tree_state["unsorted"]:
                self.tree_state["unsorted"].remove(did)
            # Close tab if open
            if did in self.tabs:
                self.tabs.remove(did)
                if self.active_tab == did:
                    self.active_tab = self.tabs[0] if self.tabs else None
            del self.tree_state["docs"][did]
            removed += 1

        if self.active_tab:
            self._switch_tab(self.active_tab, push_nav=False)
        elif self.tabs:
            self._switch_tab(self.tabs[0], push_nav=False)
        else:
            self.editor.delete("1.0", "end")

        self._render_tabs()
        self._refresh_tree()
        save_tree_state(self.tree_state)
        self._notify(f"Removed {removed} isolated wildcard(s) from explorer.", "success")

    # ── SEARCH / REPLACE IN ALL FILES ─────────────────────────────────────────
    def _show_search_all(self):
        """Search and optionally replace across all docs in the explorer."""
        win = tk.Toplevel(self.root)
        win.title("Search / Replace in All Files")
        win.configure(bg=COLORS["bg2"])
        win.resizable(True, True)
        x = self.root.winfo_x() + 60
        y = self.root.winfo_y() + 60
        win.geometry(f"680x560+{x}+{y}")

        # ── Input row ─────────────────────────────────────────────────────────
        inp = tk.Frame(win, bg=COLORS["bg2"])
        inp.pack(fill="x", padx=14, pady=(12,4))

        tk.Label(inp, text="Search:", bg=COLORS["bg2"], fg=COLORS["text1"],
                 font=("Segoe UI", 9), width=8, anchor="e").grid(row=0, column=0, padx=(0,4), pady=3)
        search_var = tk.StringVar()
        se = tk.Entry(inp, textvariable=search_var, bg=COLORS["bg3"], fg=COLORS["text0"],
                      insertbackground=COLORS["accent"], relief="flat",
                      font=("Consolas", 11), width=44)
        se.grid(row=0, column=1, sticky="ew", pady=3)

        tk.Label(inp, text="Replace:", bg=COLORS["bg2"], fg=COLORS["text1"],
                 font=("Segoe UI", 9), width=8, anchor="e").grid(row=1, column=0, padx=(0,4), pady=3)
        replace_var = tk.StringVar()
        re_entry = tk.Entry(inp, textvariable=replace_var, bg=COLORS["bg3"], fg=COLORS["text0"],
                            insertbackground=COLORS["accent"], relief="flat",
                            font=("Consolas", 11), width=44)
        re_entry.grid(row=1, column=1, sticky="ew", pady=3)
        inp.columnconfigure(1, weight=1)

        opts = tk.Frame(win, bg=COLORS["bg2"])
        opts.pack(fill="x", padx=14, pady=(0,6))
        case_var = tk.BooleanVar(value=False)
        regex_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts, text="Case sensitive", variable=case_var,
                       bg=COLORS["bg2"], fg=COLORS["text1"], selectcolor=COLORS["bg3"],
                       activebackground=COLORS["bg2"], font=("Segoe UI", 9)).pack(side="left")
        tk.Checkbutton(opts, text="Regex", variable=regex_var,
                       bg=COLORS["bg2"], fg=COLORS["text1"], selectcolor=COLORS["bg3"],
                       activebackground=COLORS["bg2"], font=("Segoe UI", 9)).pack(side="left", padx=10)

        # ── Results list ──────────────────────────────────────────────────────
        res_frame = tk.Frame(win, bg=COLORS["bg1"])
        res_frame.pack(fill="both", expand=True, padx=14, pady=(0,4))
        rsb = ttk.Scrollbar(res_frame)
        rsb.pack(side="right", fill="y")
        result_lb = tk.Listbox(res_frame, bg=COLORS["bg1"], fg=COLORS["text0"],
                               selectbackground=COLORS["sel_bg"],
                               font=("Consolas", 9), relief="flat", bd=0,
                               activestyle="none", yscrollcommand=rsb.set)
        result_lb.pack(fill="both", expand=True)
        rsb.config(command=result_lb.yview)

        status_lbl = tk.Label(win, text="", bg=COLORS["bg2"], fg=COLORS["text2"],
                              font=("Segoe UI", 9))
        status_lbl.pack(anchor="w", padx=14)

        # match_data: list of (doc_id, line_no, line_text, match_start_in_line)
        match_data = []

        def build_pattern():
            q = search_var.get()
            if not q:
                return None
            flags = 0 if case_var.get() else re.IGNORECASE
            try:
                if regex_var.get():
                    return re.compile(q, flags)
                else:
                    return re.compile(re.escape(q), flags)
            except re.error as e:
                status_lbl.config(text=f"Regex error: {e}", fg=COLORS["danger"])
                return None

        def do_search():
            match_data.clear()
            result_lb.delete(0, "end")
            pat = build_pattern()
            if not pat:
                return

            for did, doc in self.tree_state["docs"].items():
                content = doc.get("content", "")
                if did == self.active_tab:
                    content = self._get_real_content()
                lines = content.split("\n")
                for lineno, line in enumerate(lines, 1):
                    for m in pat.finditer(line):
                        match_data.append((did, lineno, line, m.start()))
                        doc_name = doc["name"]
                        result_lb.insert("end", f"  {doc_name}  :  {lineno}  |  {line[:90]}")
            status_lbl.config(
                text=f"{len(match_data)} match(es) in {len({d for d,*_ in match_data})} file(s)",
                fg=COLORS["accent3"] if match_data else COLORS["warn"])

        def open_match(event=None):
            sel = result_lb.curselection()
            if not sel:
                return
            did, lineno, _, _ = match_data[sel[0]]
            if did not in self.tabs:
                self.tabs.append(did)
            self._switch_tab(did)
            self.root.after(80, lambda: (
                self.editor.mark_set("insert", f"{lineno}.0"),
                self.editor.see(f"{lineno}.0")))

        def do_replace_all():
            pat = build_pattern()
            if not pat:
                return
            repl    = replace_var.get()
            updated = 0
            errors  = []

            for did, doc in self.tree_state["docs"].items():
                # Get the authoritative current content for this doc:
                # - active tab  → live editor text
                # - open tab    → top of its snapshot stack (most recent edit)
                # - closed tab  → doc["content"] (last synced on tab-away or save)
                if did == self.active_tab:
                    content = self._get_real_content()
                elif did in self.tabs:
                    store = self._undo_store.get(did)
                    if store and store.get("stack"):
                        content = store["stack"][store["pos"]][0]
                    else:
                        content = doc.get("content", "")
                else:
                    content = doc.get("content", "")

                new_content, count = pat.subn(repl, content)
                if count == 0:
                    continue

                # Update in-memory content
                doc["content"] = new_content
                doc["modified_unsaved"] = True
                doc["saved_hash"] = None

                if did == self.active_tab:
                    # Update the live editor and snapshot stack
                    self._undo_inhibit = True
                    try:
                        self._remove_scroll_padding()
                        self.editor.config(undo=False)
                        self.editor.delete("1.0", "end")
                        self.editor.insert("1.0", new_content)
                        self.editor.config(undo=True)
                        self.editor.edit_modified(False)
                        self._apply_scroll_padding()
                    finally:
                        self._undo_inhibit = False
                    self._push_undo_snapshot(did)
                    self._apply_wildcard_highlights()
                elif did in self.tabs:
                    # Update the snapshot stack so the tab loads correctly when switched to
                    store = self._undo_store.get(did)
                    if store and store.get("stack"):
                        pos = store["pos"]
                        # Truncate forward history then append new state
                        store["stack"] = store["stack"][:pos + 1]
                        store["stack"].append((new_content, store.get("cursor", "1.0")))
                        store["pos"] = len(store["stack"]) - 1

                # Save to disk if file has a path
                if doc.get("path"):
                    try:
                        with open(doc["path"], "w", encoding="utf-8") as fh:
                            fh.write(new_content)
                        doc["saved_hash"] = self._content_hash(new_content)
                        doc["modified_unsaved"] = False
                        updated += 1
                    except Exception as e:
                        errors.append(f"{doc['name']}: {e}")
                else:
                    updated += 1

            save_tree_state(self.tree_state)
            self._render_tabs()
            self._refresh_tree()
            msg = f"Replaced in {updated} file(s)."
            if errors:
                msg += "  Errors: " + "; ".join(errors[:3])
            status_lbl.config(text=msg,
                               fg=COLORS["accent3"] if not errors else COLORS["warn"])
            do_search()  # refresh results to confirm replacements took effect

        result_lb.bind("<Double-Button-1>", open_match)
        se.bind("<Return>", lambda e: do_search())

        btn_row = tk.Frame(win, bg=COLORS["bg2"])
        btn_row.pack(fill="x", padx=14, pady=(4, 12))
        tk.Button(btn_row, text="Search", bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=5, font=("Segoe UI", 10, "bold"),
                  command=do_search).pack(side="left", padx=(0,6))
        tk.Button(btn_row, text="Replace All", bg=COLORS["accent2"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=5, font=("Segoe UI", 10, "bold"),
                  command=do_replace_all).pack(side="left", padx=(0,6))
        tk.Button(btn_row, text="Open Selected", bg=COLORS["bg3"], fg=COLORS["text0"],
                  relief="flat", padx=10, pady=5,
                  command=open_match).pack(side="left", padx=(0,6))
        tk.Button(btn_row, text="Close", bg=COLORS["bg4"], fg=COLORS["text1"],
                  relief="flat", padx=10, pady=5,
                  command=win.destroy).pack(side="right")

        se.focus_set()

    # ── WILDCARD DIAGNOSTICS ─────────────────────────────────────────────────
    def _show_diagnostics(self):
        """Scan all explorer docs for wildcard syntax issues that would break
        dynamic-prompts / sd-wildcard parsers, and show a navigable report."""
        wrap    = self.cfg.get("wrap_str", "__")
        wc_dir  = self.cfg.get("wc_dir",   "")
        esc     = re.escape(wrap)
        ch      = wrap[0]       # e.g. '~'
        w       = len(wrap)     # e.g. 2

        # ── Pattern library ───────────────────────────────────────────────────
        # 1. Stray wrapper chars: run of ch whose length is not a multiple of w
        stray_pat = re.compile(re.escape(ch) + "+")

        # 2. Empty wildcard wrapper  e.g. ~~~~  (wrap immediately followed by wrap)
        empty_pat = re.compile(esc + esc)

        # 3. Wildcard name contains a space  e.g. ~~my wildcard~~
        space_pat = re.compile(esc + r"[^\S\n]+" + esc + "|" + esc + r"[^" + re.escape(ch) + r"\n]*\s[^" + re.escape(ch) + r"\n]*" + esc)

        # 4. Unclosed wrapper (odd number of wrap occurrences on a line)
        # We count non-overlapping occurrences of wrap on each line
        wrap_find = re.compile(esc)

        issues = []   # list of dicts: {doc_id, doc_name, line_no, col, kind, excerpt}

        # Build set of all known wildcard names (lowercased) for dead-end check
        all_doc_names = {doc["name"].lower() for doc in self.tree_state["docs"].values()}
        ref_pattern = re.compile(esc + r"([^\s]+?)" + esc)

        def scan_doc(doc_id, doc_name, content):
            lines = content.split("\n")
            for lineno, line in enumerate(lines, 1):
                # ── Check 1: stray wrapper characters ─────────────────────────
                for m in stray_pat.finditer(line):
                    run = m.group()
                    if len(run) % w != 0:
                        issues.append({
                            "doc_id":   doc_id,
                            "doc_name": doc_name,
                            "line_no":  lineno,
                            "col":      m.start() + 1,
                            "kind":     "Stray wrapper char",
                            "excerpt":  _excerpt(line, m.start()),
                            "fix":      "stray",
                            "m_start":  m.start(),
                            "m_end":    m.end(),
                        })

                # ── Check 2: empty wildcard ────────────────────────────────────
                for m in empty_pat.finditer(line):
                    issues.append({
                        "doc_id":   doc_id,
                        "doc_name": doc_name,
                        "line_no":  lineno,
                        "col":      m.start() + 1,
                        "kind":     "Empty wildcard wrapper",
                        "excerpt":  _excerpt(line, m.start()),
                        "fix":      None,
                        "m_start":  m.start(),
                        "m_end":    m.end(),
                    })

                # ── Check 3: unclosed wrapper on this line ─────────────────────
                count = len(wrap_find.findall(line))
                if count % 2 != 0:
                    positions = [m.start() for m in wrap_find.finditer(line)]
                    col = positions[-1] + 1
                    issues.append({
                        "doc_id":   doc_id,
                        "doc_name": doc_name,
                        "line_no":  lineno,
                        "col":      col,
                        "kind":     f"Unclosed wrapper ({count} occurrence{'s' if count!=1 else ''})",
                        "excerpt":  _excerpt(line, col - 1),
                        "fix":      None,
                        "m_start":  col - 1,
                        "m_end":    col - 1 + w,
                    })

                # ── Check 4: dead-end wildcard (called but not in explorer) ────
                for m in ref_pattern.finditer(line):
                    ref_name = m.group(1).lower()
                    if ref_name not in all_doc_names:
                        issues.append({
                            "doc_id":   doc_id,
                            "doc_name": doc_name,
                            "line_no":  lineno,
                            "col":      m.start() + 1,
                            "kind":     f"Dead-end: __{m.group(1)}__ not found",
                            "excerpt":  _excerpt(line, m.start()),
                            "fix":      None,
                            "m_start":  m.start(),
                            "m_end":    m.end(),
                        })

        def _excerpt(line, col, radius=40):
            start = max(0, col - radius)
            end   = min(len(line), col + radius)
            snip  = line[start:end].strip()
            return snip[:80]

        # Scan all tracked docs (use live editor content for active tab)
        for did, doc in self.tree_state["docs"].items():
            content = doc.get("content", "")
            if did == self.active_tab:
                content = self._get_real_content()
            if content.strip():
                scan_doc(did, doc["name"], content)

        # ── Build window ──────────────────────────────────────────────────────
        win = tk.Toplevel(self.root)
        win.title("Wildcard Diagnostics")
        win.configure(bg=COLORS["bg2"])
        win.resizable(True, True)
        x = self.root.winfo_x() + 50
        y = self.root.winfo_y() + 40
        win.geometry(f"820x560+{x}+{y}")

        # Header
        hdr = tk.Frame(win, bg=COLORS["bg2"])
        hdr.pack(fill="x", padx=14, pady=(12, 4))
        if issues:
            tk.Label(hdr, text=f"⚠  {len(issues)} issue(s) found across {len({i['doc_id'] for i in issues})} file(s)",
                     bg=COLORS["bg2"], fg=COLORS["warn"],
                     font=("Segoe UI", 11, "bold")).pack(side="left")
        else:
            tk.Label(hdr, text="✓  No wildcard syntax issues found.",
                     bg=COLORS["bg2"], fg=COLORS["accent3"],
                     font=("Segoe UI", 11, "bold")).pack(side="left")
            tk.Button(hdr, text="Close", bg=COLORS["accent"], fg=COLORS["bg0"],
                      relief="flat", padx=12, pady=4,
                      command=win.destroy).pack(side="right")
            return

        tk.Label(win,
                 text="Double-click any row to jump to that file and line.  "
                      "Stray chars can be auto-fixed with the Fix button.",
                 bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9), wraplength=790).pack(anchor="w", padx=14, pady=(0,6))

        # Filter row
        filt_frame = tk.Frame(win, bg=COLORS["bg2"])
        filt_frame.pack(fill="x", padx=14, pady=(0, 4))
        tk.Label(filt_frame, text="Filter file:", bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9)).pack(side="left")
        filter_var = tk.StringVar()
        filt_entry = tk.Entry(filt_frame, textvariable=filter_var,
                              bg=COLORS["bg3"], fg=COLORS["text0"],
                              insertbackground=COLORS["accent"],
                              relief="flat", font=("Consolas", 10), width=28)
        filt_entry.pack(side="left", padx=6)

        # Results table (Treeview)
        cols = ("file", "line", "col", "kind", "excerpt")
        tree_frame = tk.Frame(win, bg=COLORS["bg1"])
        tree_frame.pack(fill="both", expand=True, padx=14, pady=(0,4))
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        vsb.pack(side="right", fill="y")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        hsb.pack(side="bottom", fill="x")

        style = ttk.Style()
        style.configure("Diag.Treeview",
                         background=COLORS["bg1"], foreground=COLORS["text0"],
                         fieldbackground=COLORS["bg1"], rowheight=22,
                         font=("Consolas", 9))
        style.configure("Diag.Treeview.Heading",
                         background=COLORS["bg3"], foreground=COLORS["text1"],
                         font=("Segoe UI", 9, "bold"))
        style.map("Diag.Treeview", background=[("selected", COLORS["sel_bg"])])

        tv = ttk.Treeview(tree_frame, columns=cols, show="headings",
                          style="Diag.Treeview",
                          yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tv.pack(fill="both", expand=True)
        vsb.config(command=tv.yview)
        hsb.config(command=tv.xview)

        tv.heading("file",    text="File")
        tv.heading("line",    text="Line")
        tv.heading("col",     text="Col")
        tv.heading("kind",    text="Issue")
        tv.heading("excerpt", text="Context")
        tv.column("file",    width=160, minwidth=80)
        tv.column("line",    width=50,  minwidth=40, anchor="center")
        tv.column("col",     width=45,  minwidth=40, anchor="center")
        tv.column("kind",    width=200, minwidth=120)
        tv.column("excerpt", width=340, minwidth=100)

        # Tag colouring for issue types
        tv.tag_configure("stray",    foreground=COLORS["danger"])
        tv.tag_configure("empty",    foreground=COLORS["warn"])
        tv.tag_configure("unclosed", foreground=COLORS["accent2"])
        tv.tag_configure("deadend",  foreground=COLORS["accent3"])

        displayed = []   # parallel list to tv rows

        def populate(filter_text=""):
            for row in tv.get_children():
                tv.delete(row)
            displayed.clear()
            ft = filter_text.lower()
            for issue in issues:
                if ft and ft not in issue["doc_name"].lower():
                    continue
                tag = ("stray"    if "Stray"    in issue["kind"] else
                       "empty"    if "Empty"    in issue["kind"] else
                       "deadend"  if "Dead-end" in issue["kind"] else
                       "unclosed")
                tv.insert("", "end",
                           values=(issue["doc_name"],
                                   issue["line_no"],
                                   issue["col"],
                                   issue["kind"],
                                   issue["excerpt"]),
                           tags=(tag,))
                displayed.append(issue)

        populate()
        filter_var.trace_add("write", lambda *_: populate(filter_var.get()))

        def jump_to_issue(event=None):
            sel = tv.selection()
            if not sel:
                return
            idx   = tv.index(sel[0])
            issue = displayed[idx]
            did   = issue["doc_id"]
            if did not in self.tabs:
                self.tabs.append(did)
            self._switch_tab(did)
            lineno = issue["line_no"]
            col    = issue["m_start"]
            self.root.after(80, lambda: (
                self.editor.mark_set("insert", f"{lineno}.{col}"),
                self.editor.see(f"{lineno}.{col}"),
                # Briefly flash the problem span
                self.editor.tag_add("warn_tilde", f"{lineno}.{issue['m_start']}", f"{lineno}.{issue['m_end']}"),
                self.editor.tag_raise("warn_tilde"),
            ))

        def fix_stray(issue):
            """Auto-fix a stray-char issue by padding or trimming the run to nearest
            valid multiple of wrap length."""
            did = issue["doc_id"]
            if did not in self.tabs:
                self.tabs.append(did)
            self._switch_tab(did)
            lineno = issue["line_no"]

            def do_fix():
                try:
                    line_start = f"{lineno}.0"
                    line_end   = f"{lineno}.end"
                    line_text  = self._get_real_content().split("\n")[lineno - 1]
                    m_start    = issue["m_start"]
                    m_end      = issue["m_end"]
                    run        = line_text[m_start:m_end]
                    run_len    = len(run)
                    # Round to nearest multiple of w
                    lower = (run_len // w) * w
                    upper = lower + w
                    target_len = lower if abs(run_len - lower) <= abs(run_len - upper) else upper
                    if target_len == 0:
                        target_len = w
                    new_run = ch * target_len
                    # Replace in editor
                    abs_start = f"{lineno}.{m_start}"
                    abs_end   = f"{lineno}.{m_end}"
                    self.editor.delete(abs_start, abs_end)
                    self.editor.insert(abs_start, new_run)
                    self._notify(f"Fixed: replaced {run!r} → {new_run!r} on line {lineno}", "success")
                except Exception as e:
                    self._notify(f"Fix failed: {e}", "warn")
            self.root.after(120, do_fix)

        def on_fix_btn():
            sel = tv.selection()
            if not sel:
                self._notify("Select a row first.", "warn")
                return
            idx   = tv.index(sel[0])
            issue = displayed[idx]
            if issue["fix"] == "stray":
                fix_stray(issue)
            else:
                self._notify("No auto-fix available for this issue type — jump to it and edit manually.", "warn")

        tv.bind("<Double-Button-1>", jump_to_issue)

        # Button row
        btn_row = tk.Frame(win, bg=COLORS["bg2"])
        btn_row.pack(fill="x", padx=14, pady=(4, 12))
        tk.Button(btn_row, text="Jump to Issue",
                  bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=5, font=("Segoe UI", 10, "bold"),
                  command=jump_to_issue).pack(side="left", padx=(0,6))
        tk.Button(btn_row, text="Auto-Fix Stray Char",
                  bg=COLORS["accent2"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=5, font=("Segoe UI", 10, "bold"),
                  command=on_fix_btn).pack(side="left", padx=(0,6))
        tk.Label(btn_row,
                 text=f"{len(issues)} issue(s)  •  {len({i['doc_id'] for i in issues})} file(s) affected",
                 bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9)).pack(side="left", padx=12)
        tk.Button(btn_row, text="Close",
                  bg=COLORS["bg4"], fg=COLORS["text1"],
                  relief="flat", padx=10, pady=5,
                  command=win.destroy).pack(side="right")

    # ── LORA STRENGTH ADJUSTER ────────────────────────────────────────────────
    def _show_lora_adjust(self):
        """Increment or decrement a LoRA's strength value across one doc or all docs."""
        wc_dir = self.cfg.get("wc_dir", "")
        win = tk.Toplevel(self.root)
        win.title("LoRA Strength Adjust")
        win.configure(bg=COLORS["bg2"])
        win.resizable(False, False)
        x = self.root.winfo_x() + 100
        y = self.root.winfo_y() + 100
        win.geometry(f"520x320+{x}+{y}")
        win.update_idletasks()
        win.grab_set()

        pad = dict(padx=14, pady=4)

        tk.Label(win, text="LoRA Strength Adjuster",
                 bg=COLORS["bg2"], fg=COLORS["accent"],
                 font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=14, pady=(12,2))

        tk.Label(win,
                 text="Enter the LoRA pattern with # as the strength placeholder.\n"
                      "Example:  <lora:my_lora:#>",
                 bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9), justify="left").pack(anchor="w", padx=14, pady=(0,6))

        row1 = tk.Frame(win, bg=COLORS["bg2"])
        row1.pack(fill="x", **pad)
        tk.Label(row1, text="Pattern:", bg=COLORS["bg2"], fg=COLORS["text1"],
                 font=("Segoe UI", 9), width=10, anchor="e").pack(side="left")
        pattern_var = tk.StringVar(value="<lora::#>")
        pattern_entry = tk.Entry(row1, textvariable=pattern_var,
                                 bg=COLORS["bg3"], fg=COLORS["text0"],
                                 insertbackground=COLORS["accent"],
                                 font=("Consolas", 10), relief="flat", width=38)
        pattern_entry.pack(side="left", padx=6)

        row2 = tk.Frame(win, bg=COLORS["bg2"])
        row2.pack(fill="x", **pad)
        tk.Label(row2, text="Increment:", bg=COLORS["bg2"], fg=COLORS["text1"],
                 font=("Segoe UI", 9), width=10, anchor="e").pack(side="left")
        incr_var = tk.StringVar(value="0.1")
        incr_entry = tk.Entry(row2, textvariable=incr_var,
                              bg=COLORS["bg3"], fg=COLORS["text0"],
                              insertbackground=COLORS["accent"],
                              font=("Consolas", 10), relief="flat", width=10)
        incr_entry.pack(side="left", padx=6)
        tk.Label(row2, text="(use negative to decrement, e.g. -0.1)",
                 bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9)).pack(side="left", padx=4)

        status_lbl = tk.Label(win, text="", bg=COLORS["bg2"], fg=COLORS["text2"],
                              font=("Segoe UI", 9))
        status_lbl.pack(anchor="w", padx=14, pady=4)

        def build_regex(pattern_str):
            """Convert user pattern (with #) into a regex that captures the number."""
            if "#" not in pattern_str:
                return None, "Pattern must contain # as the strength placeholder."
            # Escape everything except #, then replace # with number capture group
            parts = pattern_str.split("#")
            regex_str = re.escape(parts[0]) + r"([-+]?\d*\.?\d+)" + re.escape(parts[1])
            try:
                return re.compile(regex_str), None
            except re.error as e:
                return None, str(e)

        def adjust_content(content, pat, incr):
            """Replace all matches of pat in content, adding incr to the captured number."""
            def replacer(m):
                try:
                    val = float(m.group(1))
                    new_val = round(val + incr, 6)
                    # Preserve sign for negative values; format cleanly
                    formatted = f"{new_val:.4g}"
                    return m.group(0).replace(m.group(1), formatted, 1)
                except Exception:
                    return m.group(0)
            new_content, count = pat.subn(replacer, content)
            return new_content, count

        def do_adjust(scope):
            pattern_str = pattern_var.get().strip()
            try:
                incr = float(incr_var.get().strip())
            except ValueError:
                status_lbl.config(text="Invalid increment value.", fg=COLORS["danger"])
                return

            pat, err = build_regex(pattern_str)
            if not pat:
                status_lbl.config(text=f"Pattern error: {err}", fg=COLORS["danger"])
                return

            if scope == "all":
                if not messagebox.askyesno("Adjust All Wildcards",
                        f"Apply {'+' if incr >= 0 else ''}{incr} to all matching LoRA strengths "
                        f"in ALL wildcard files?\n\nThis will save all affected files.",
                        parent=win):
                    return

            total_changes = 0
            total_files   = 0

            docs_to_process = {}
            if scope == "current":
                if not self.active_tab:
                    status_lbl.config(text="No active document.", fg=COLORS["danger"])
                    return
                docs_to_process[self.active_tab] = self.tree_state["docs"][self.active_tab]
            else:
                docs_to_process = dict(self.tree_state["docs"])

            for did, doc in docs_to_process.items():
                content = doc.get("content", "")
                if did == self.active_tab:
                    content = self._get_real_content()
                new_content, count = adjust_content(content, pat, incr)
                if count == 0:
                    continue
                total_changes += count
                total_files   += 1
                doc["content"] = new_content
                doc["modified_unsaved"] = True
                if did == self.active_tab:
                    self._push_undo_snapshot()
                    self.editor.config(undo=False)
                    self._remove_scroll_padding()
                    self.editor.delete("1.0", "end")
                    self.editor.insert("1.0", new_content)
                    self.editor.config(undo=True)
                    self.editor.edit_modified(False)
                    self._apply_scroll_padding()
                    self._apply_wildcard_highlights()
                    self._push_undo_snapshot()
                if doc.get("path") and scope == "all":
                    try:
                        with open(doc["path"], "w", encoding="utf-8") as fh:
                            fh.write(new_content)
                        doc["saved_hash"] = self._content_hash(new_content)
                        doc["modified_unsaved"] = False
                    except Exception as e:
                        status_lbl.config(text=f"Save error: {e}", fg=COLORS["danger"])

            save_tree_state(self.tree_state)
            self._render_tabs()
            self._refresh_tree()
            msg = (f"Changed {total_changes} value(s) across {total_files} file(s)."
                   if total_files else "No matching LoRA patterns found.")
            status_lbl.config(text=msg, fg=COLORS["accent3"] if total_files else COLORS["warn"])

        btn_row = tk.Frame(win, bg=COLORS["bg2"])
        btn_row.pack(fill="x", padx=14, pady=(8, 14))
        tk.Button(btn_row, text="Adjust This Document",
                  bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=6,
                  font=("Segoe UI", 10, "bold"),
                  command=lambda: do_adjust("current")).pack(side="left", padx=(0,8))
        tk.Button(btn_row, text="Adjust All Wildcards",
                  bg=COLORS["accent2"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=6,
                  font=("Segoe UI", 10, "bold"),
                  command=lambda: do_adjust("all")).pack(side="left", padx=(0,8))
        tk.Button(btn_row, text="Close",
                  bg=COLORS["bg4"], fg=COLORS["text1"],
                  relief="flat", padx=10, pady=6,
                  command=win.destroy).pack(side="right")

        pattern_entry.focus_set()
        pattern_entry.icursor("end")

    def _show_reorg_confirm(self):
        win = tk.Toplevel(self.root)
        win.title("Reorganize Files")
        win.configure(bg=COLORS["bg2"])
        win.resizable(False, False)
        x = self.root.winfo_x() + 150
        y = self.root.winfo_y() + 150
        win.geometry(f"400x220+{x}+{y}")
        win.update_idletasks()
        win.grab_set()
        tk.Label(win, text="⚠ Reorganize Files on Disk",
                 bg=COLORS["bg2"], fg=COLORS["warn"],
                 font=("Segoe UI", 12, "bold")).pack(pady=(16,8), padx=16)
        tk.Label(win,
                 text="This will physically move files on disk to match\n"
                      "the folder structure shown in the sidebar.",
                 bg=COLORS["bg2"], fg=COLORS["text1"],
                 font=("Segoe UI", 10), justify="center").pack(pady=4)
        tk.Label(win, text="⚠ This cannot be undone. Back up your files first.",
                 bg=COLORS["bg2"], fg=COLORS["danger"],
                 font=("Segoe UI", 9)).pack(pady=4)
        btn_f = tk.Frame(win, bg=COLORS["bg2"])
        btn_f.pack(pady=12)
        tk.Button(btn_f, text="Cancel", bg=COLORS["bg3"], fg=COLORS["text1"],
                  relief="flat", padx=12, pady=6, command=win.destroy).pack(side="left", padx=6)
        tk.Button(btn_f, text="Yes, Reorganize", bg=COLORS["danger"], fg="white",
                  relief="flat", padx=12, pady=6,
                  command=lambda: [self._do_reorganize(), win.destroy()]).pack(side="left", padx=6)

    def _do_reorganize(self):
        wc_dir = self.cfg.get("wc_dir","")
        if not wc_dir or not os.path.isdir(wc_dir):
            messagebox.showerror("Error", f"Wildcards directory not found:\n{wc_dir}\n\nPlease update in Settings.")
            return

        moved  = 0
        errors = []

        # ── Build folder-id → parent-path map ────────────────────────────────
        # Identify top-level folders (not referenced as anyone's child)
        child_folder_ids = set()
        for f in self.tree_state["folders"]:
            for cid in f.get("children", []):
                child_folder_ids.add(cid)

        folder_by_id = {f["id"]: f for f in self.tree_state["folders"]}

        # Recursively compute the absolute disk path for a folder by walking
        # the parent chain. We build a cache to avoid repeated traversal.
        folder_path_cache = {}

        def get_folder_path(folder_id):
            if folder_id in folder_path_cache:
                return folder_path_cache[folder_id]
            folder = folder_by_id.get(folder_id)
            if not folder:
                return wc_dir
            # Find parent of this folder (if any)
            parent = None
            for f in self.tree_state["folders"]:
                if folder_id in f.get("children", []):
                    parent = f
                    break
            if parent is None:
                # Top-level folder → lives directly in wc_dir
                path = os.path.join(wc_dir, folder["name"])
            else:
                # Subfolder → lives inside parent's path
                parent_path = get_folder_path(parent["id"])
                path = os.path.join(parent_path, folder["name"])
            folder_path_cache[folder_id] = path
            return path

        # Pre-compute paths for all folders
        for f in self.tree_state["folders"]:
            get_folder_path(f["id"])

        # ── Process each folder recursively (top-down so dirs exist before move) ─
        def process_folder(folder):
            folder_path = get_folder_path(folder["id"])
            os.makedirs(folder_path, exist_ok=True)
            # Move docs into this folder
            for doc_id in folder.get("docs", []):
                doc = self.tree_state["docs"].get(doc_id)
                if not doc or not doc.get("path"):
                    continue
                src = doc["path"]
                dst = os.path.join(folder_path, Path(src).name)
                if os.path.normpath(src) == os.path.normpath(dst):
                    continue
                try:
                    if os.path.exists(src):
                        shutil.move(src, dst)
                        doc["path"] = dst
                        moved += 1
                except Exception as e:
                    errors.append(f"{Path(src).name}: {e}")
            # Recurse into child folders first
            for cid in folder.get("children", []):
                child = folder_by_id.get(cid)
                if child:
                    process_folder(child)

        top_folders = [f for f in self.tree_state["folders"]
                       if f["id"] not in child_folder_ids]
        for f in top_folders:
            process_folder(f)

        # ── Delete any now-empty directories under wc_dir ─────────────────────
        # Walk bottom-up so deepest dirs are evaluated first
        deleted_dirs = []
        for dirpath, dirnames, filenames in os.walk(wc_dir, topdown=False):
            if os.path.normpath(dirpath) == os.path.normpath(wc_dir):
                continue  # never delete the root wc_dir itself
            try:
                # A directory is empty if it has no files and no subdirs
                if not os.listdir(dirpath):
                    os.rmdir(dirpath)
                    deleted_dirs.append(dirpath)
            except Exception:
                pass

        save_tree_state(self.tree_state)

        msg = f"Moved {moved} file(s)."
        if deleted_dirs:
            msg += f"\nDeleted {len(deleted_dirs)} empty folder(s)."
        if errors:
            msg += f"\n\nErrors ({len(errors)}):\n" + "\n".join(errors[:5])
        self._notify(msg.split("\n")[0], "success" if not errors else "warn")
        messagebox.showinfo("Reorganize Complete", msg)

    # ── SETTINGS ─────────────────────────────────────────────────────────────
    def _show_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Settings")
        win.configure(bg=COLORS["bg2"])
        win.resizable(False, False)
        x = self.root.winfo_x() + 100
        y = self.root.winfo_y() + 80
        win.geometry(f"520x460+{x}+{y}")
        win.update_idletasks()
        win.grab_set()

        def row(parent, label, widget_fn, row_n):
            tk.Label(parent, text=label, bg=COLORS["bg2"], fg=COLORS["text2"],
                     font=("Segoe UI", 9, "bold"), anchor="w").grid(
                     row=row_n, column=0, padx=(16,8), pady=6, sticky="w")
            w = widget_fn(parent)
            w.grid(row=row_n, column=1, padx=(0,16), pady=6, sticky="ew")
            return w

        frame = tk.Frame(win, bg=COLORS["bg2"])
        frame.pack(fill="both", expand=True, padx=0, pady=0)
        frame.columnconfigure(1, weight=1)

        tk.Label(frame, text="Settings", bg=COLORS["bg2"], fg=COLORS["text0"],
                 font=("Segoe UI", 13, "bold")).grid(
                 row=0, column=0, columnspan=2, padx=16, pady=(16,8), sticky="w")

        def entry_w(parent, val):
            e = tk.Entry(parent, bg=COLORS["bg3"], fg=COLORS["text0"],
                         insertbackground=COLORS["accent"],
                         relief="flat", bd=1, font=("Consolas", 10),
                         highlightthickness=1,
                         highlightcolor=COLORS["accent"],
                         highlightbackground=COLORS["border2"])
            e.insert(0, val)
            return e

        wrap_entry = row(frame, "Wildcard Wrap String",
                          lambda p: entry_w(p, self.cfg["wrap_str"]), 1)
        dir_entry  = row(frame, "Wildcards Directory",
                          lambda p: entry_w(p, self.cfg["wc_dir"]), 2)

        def browse_dir():
            d = filedialog.askdirectory(initialdir=self.cfg.get("wc_dir",""))
            if d:
                dir_entry.delete(0,"end")
                dir_entry.insert(0,d)
        browse_btn = tk.Label(frame, text="Browse…", bg=COLORS["bg3"],
                               fg=COLORS["accent"], font=("Segoe UI",9),
                               cursor="hand2", padx=6)
        browse_btn.grid(row=2, column=2, padx=(0,16), pady=6)
        browse_btn.bind("<Button-1>", lambda e: browse_dir())

        fs_var = tk.IntVar(value=self.cfg["font_size"])
        fs_entry = row(frame, "Editor Font Size",
                        lambda p: tk.Spinbox(p, from_=9, to=24, textvariable=fs_var,
                                             bg=COLORS["bg3"], fg=COLORS["text0"],
                                             buttonbackground=COLORS["bg4"],
                                             relief="flat", width=5,
                                             font=("Consolas",10),
                                             highlightthickness=0), 3)

        # Font family picker
        FONT_OPTIONS = [
            "Consolas", "Courier New", "Cascadia Code", "Cascadia Mono",
            "Fira Code", "JetBrains Mono", "Source Code Pro", "Lucida Console",
            "Inconsolata", "Monaco",
        ]
        current_font = self.cfg.get("font_family", "Consolas")
        font_var = tk.StringVar(value=current_font)
        tk.Label(frame, text="Editor Font", bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9, "bold"), anchor="w").grid(
                 row=4, column=0, padx=(16,8), pady=6, sticky="w")
        # Style the combobox so the collapsed entry reads black-on-white
        style = ttk.Style()
        style.configure("FontPicker.TCombobox",
                         fieldbackground="white",
                         background="white",
                         foreground="black",
                         selectbackground="white",
                         selectforeground="black")
        font_menu = ttk.Combobox(frame, textvariable=font_var, values=FONT_OPTIONS,
                                  state="readonly", width=22, style="FontPicker.TCombobox")
        font_menu.grid(row=4, column=1, padx=(0,16), pady=6, sticky="ew")

        spell_var = tk.BooleanVar(value=self.cfg["spell_check"])
        tk.Checkbutton(frame, text="Enable Spell Check",
                        variable=spell_var,
                        bg=COLORS["bg2"], fg=COLORS["text1"],
                        selectcolor=COLORS["bg3"],
                        activebackground=COLORS["bg2"],
                        font=("Segoe UI",10)).grid(
                        row=5, column=0, columnspan=2, padx=16, pady=4, sticky="w")

        auto_var = tk.BooleanVar(value=self.cfg.get("autosave",False))
        tk.Checkbutton(frame, text="Auto-save on tab switch",
                        variable=auto_var,
                        bg=COLORS["bg2"], fg=COLORS["text1"],
                        selectcolor=COLORS["bg3"],
                        activebackground=COLORS["bg2"],
                        font=("Segoe UI",10)).grid(
                        row=6, column=0, columnspan=2, padx=16, pady=4, sticky="w")

        # User dictionary editor link
        dict_path = CONFIG_PATH.parent / "user_dict.txt"
        dict_info = tk.Label(frame,
                             text=f"User Dictionary:  {dict_path}",
                             bg=COLORS["bg2"], fg=COLORS["text2"],
                             font=("Segoe UI", 8), anchor="w")
        dict_info.grid(row=7, column=0, columnspan=2, padx=(16,0), pady=(6,0), sticky="w")

        edit_dict_btn = tk.Label(frame, text="Edit Added Words…",
                                  bg=COLORS["bg3"], fg=COLORS["accent"],
                                  font=("Segoe UI", 9), cursor="hand2", padx=8, pady=3)
        edit_dict_btn.grid(row=7, column=2, padx=(0,16), pady=(6,0))
        edit_dict_btn.bind("<Button-1>", lambda e: self._show_user_dict_editor(win))

        def save_s():
            self.cfg["wrap_str"]     = wrap_entry.get() or "__"
            self.cfg["wc_dir"]       = dir_entry.get()
            self.cfg["font_size"]    = fs_var.get()
            self.cfg["font_family"]  = font_var.get()
            self.cfg["spell_check"]  = spell_var.get()
            self.cfg["autosave"]     = auto_var.get()
            self.spell_enabled       = spell_var.get()
            fam = self.cfg["font_family"]
            sz  = self.cfg["font_size"]
            self.editor.config(font=(fam, sz))
            # Keep bracket_match bold tag in sync with font settings
            self.editor.tag_configure("bracket_match",
                                       font=(fam, sz, "bold"),
                                       foreground="#ffffff")
            self._apply_scroll_padding()
            self._redraw_line_numbers()
            self.sb_wrap_str.config(text=f"Wrap: {self.cfg['wrap_str']}")
            self.sb_spell.config(text="Spell ✓" if self.spell_enabled else "Spell ✗",
                                  fg=COLORS["accent3"] if self.spell_enabled else COLORS["text2"])
            save_config(self.cfg)
            self._apply_wildcard_highlights()
            self._update_wc_list()
            win.destroy()
            self._notify("Settings saved", "success")

        btn_f = tk.Frame(frame, bg=COLORS["bg2"])
        btn_f.grid(row=8, column=0, columnspan=3, pady=16, padx=16, sticky="e")
        tk.Button(btn_f, text="Cancel", bg=COLORS["bg3"], fg=COLORS["text1"],
                  relief="flat", padx=12, pady=6, command=win.destroy).pack(side="left", padx=4)
        tk.Button(btn_f, text="Save Settings", bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=12, pady=6, font=("Segoe UI",10,"bold"),
                  command=save_s).pack(side="left", padx=4)

    def _show_user_dict_editor(self, parent_win=None):
        """Show a window listing custom-added dictionary words with add/remove."""
        dict_path = CONFIG_PATH.parent / "user_dict.txt"

        win = tk.Toplevel(parent_win or self.root)
        win.title("User Dictionary")
        win.configure(bg=COLORS["bg2"])
        win.resizable(True, True)
        x = self.root.winfo_x() + 80
        y = self.root.winfo_y() + 100
        win.geometry(f"340x480+{x}+{y}")
        win.update_idletasks()
        win.grab_set()

        tk.Label(win, text="User Dictionary",
                 bg=COLORS["bg2"], fg=COLORS["text0"],
                 font=("Segoe UI", 12, "bold")).pack(pady=(14,2), padx=16, anchor="w")
        tk.Label(win, text="Select a word and press Remove to delete it.",
                 bg=COLORS["bg2"], fg=COLORS["text2"],
                 font=("Segoe UI", 9)).pack(padx=16, anchor="w")

        list_frame = tk.Frame(win, bg=COLORS["bg1"], relief="flat", bd=1)
        list_frame.pack(fill="both", expand=True, padx=16, pady=(8,4))

        scroll = ttk.Scrollbar(list_frame, orient="vertical")
        scroll.pack(side="right", fill="y")

        lb = tk.Listbox(list_frame,
                         bg=COLORS["bg1"], fg=COLORS["text0"],
                         selectbackground=COLORS["sel_bg"],
                         selectforeground=COLORS["text0"],
                         font=("Consolas", 11),
                         relief="flat", bd=0,
                         activestyle="none",
                         yscrollcommand=scroll.set)
        lb.pack(side="left", fill="both", expand=True)
        scroll.config(command=lb.yview)

        def reload_words():
            lb.delete(0, "end")
            if dict_path.exists():
                try:
                    words = [w.strip() for w in dict_path.read_text(encoding="utf-8").splitlines() if w.strip()]
                    for w in sorted(set(words)):
                        lb.insert("end", f"  {w}")
                except Exception:
                    pass
            if lb.size() == 0:
                lb.insert("end", "  (no custom words)")
                lb.config(state="disabled")
            else:
                lb.config(state="normal")

        reload_words()

        def remove_selected():
            sel = lb.curselection()
            if not sel:
                return
            word = lb.get(sel[0]).strip()
            if not word or word.startswith("("):
                return
            ans = messagebox.askyesno(
                "Remove Word",
                f'Remove "{word}" from your dictionary?',
                parent=win)
            if not ans:
                return
            try:
                if dict_path.exists():
                    lines = [l for l in dict_path.read_text(encoding="utf-8").splitlines()
                             if l.strip().lower() != word.lower()]
                    dict_path.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")
                if hasattr(self, "_spell_instance"):
                    del self._spell_instance
                reload_words()
                self._run_spell_check()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=win)

        # ── Add word section ──────────────────────────────────────────────────
        add_frame = tk.Frame(win, bg=COLORS["bg2"])
        add_frame.pack(fill="x", padx=16, pady=(4,2))

        tk.Label(add_frame, text="Add word:", bg=COLORS["bg2"], fg=COLORS["text1"],
                 font=("Segoe UI", 9)).pack(side="left")

        add_var = tk.StringVar()
        add_entry = tk.Entry(add_frame, textvariable=add_var,
                              bg=COLORS["bg3"], fg=COLORS["text0"],
                              insertbackground=COLORS["accent"],
                              relief="flat", bd=1, font=("Consolas", 10),
                              highlightthickness=1,
                              highlightcolor=COLORS["accent"],
                              highlightbackground=COLORS["border2"])
        add_entry.pack(side="left", fill="x", expand=True, padx=(6,6))

        def do_add_word(event=None):
            word = add_var.get().strip().lower()
            if not word or not word.isalpha():
                return
            try:
                existing = []
                if dict_path.exists():
                    existing = [l.strip() for l in dict_path.read_text(encoding="utf-8").splitlines() if l.strip()]
                if word not in existing:
                    with open(dict_path, "a", encoding="utf-8") as f:
                        f.write(word + "\n")
                if hasattr(self, "_spell_instance"):
                    del self._spell_instance
                add_var.set("")
                reload_words()
                self._run_spell_check()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=win)

        add_entry.bind("<Return>", do_add_word)
        tk.Button(add_frame, text="Add", bg=COLORS["accent3"], fg=COLORS["bg0"],
                  relief="flat", padx=8, pady=2, font=("Segoe UI", 9, "bold"),
                  command=do_add_word).pack(side="left")

        # ── Bottom buttons ────────────────────────────────────────────────────
        btn_row = tk.Frame(win, bg=COLORS["bg2"])
        btn_row.pack(fill="x", padx=16, pady=(4, 14))
        tk.Button(btn_row, text="Remove Selected", bg=COLORS["danger"], fg=COLORS["bg0"],
                  relief="flat", padx=10, pady=5, font=("Segoe UI", 9),
                  command=remove_selected).pack(side="left")
        tk.Button(btn_row, text="Close", bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=14, pady=5, font=("Segoe UI", 10, "bold"),
                  command=win.destroy).pack(side="right")

    # ── HOTKEYS MODAL ────────────────────────────────────────────────────────
    def _show_hotkeys(self):
        win = tk.Toplevel(self.root)
        win.title("Keyboard Shortcuts")
        win.configure(bg=COLORS["bg2"])
        win.resizable(False, False)
        x = self.root.winfo_x() + 120
        y = self.root.winfo_y() + 80
        win.geometry(f"420x500+{x}+{y}")
        win.update_idletasks()
        win.grab_set()

        tk.Label(win, text="Keyboard Shortcuts", bg=COLORS["bg2"], fg=COLORS["text0"],
                 font=("Segoe UI",13,"bold")).pack(pady=(16,8), padx=16, anchor="w")

        frame = tk.Frame(win, bg=COLORS["bg1"])
        frame.pack(fill="both", expand=True, padx=16, pady=(0,8))

        shortcuts = [
            ("New File",          "Ctrl+N"),
            ("Open File",         "Ctrl+O"),
            ("Save",              "Ctrl+S"),
            ("Save As",           "Ctrl+Shift+S"),
            ("Undo",              "Ctrl+Z"),
            ("Redo",              "Ctrl+Y"),
            ("Clone Lines",       "Ctrl+D"),
            ("Wrap as Wildcard",  "Ctrl+Shift+W"),
            ("Find (with selection)", "Ctrl+F  /  Ctrl+H"),
            ("Search All Files",    "toolbar button"),
            ("Find Next",         "F3"),
            ("Find Previous",     "Shift+F3"),
            ("Nav Back",          "Alt+Left"),
            ("Nav Forward",       "Alt+Right"),
            ("Close Tab",         "Ctrl+W  /  Ctrl+F4"),
            ("Toggle Spell",      "F7"),
            ("Toggle Word Wrap",  "Ctrl+Shift+Z"),
            ("New Folder",        "Ctrl+Shift+N"),
            ("Settings",          "Ctrl+,"),
            ("Hotkeys",           "Ctrl+/"),
        ]

        for i, (action, shortcut) in enumerate(shortcuts):
            bg = COLORS["bg1"] if i%2==0 else COLORS["bg2"]
            row = tk.Frame(frame, bg=bg)
            row.pack(fill="x")
            tk.Label(row, text=action, bg=bg, fg=COLORS["text1"],
                     font=("Segoe UI",10), anchor="w", padx=12, pady=5).pack(side="left")
            tk.Label(row, text=shortcut, bg=bg, fg=COLORS["accent"],
                     font=("Consolas",10), anchor="e", padx=12).pack(side="right")

        tk.Button(win, text="Close", bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=16, pady=6, font=("Segoe UI",10,"bold"),
                  command=win.destroy).pack(pady=12)

    # ── NOTIFICATION ──────────────────────────────────────────────────────────
    def _notify(self, msg, kind="info"):
        colors = {"info": COLORS["accent"], "success": COLORS["accent3"],
                  "warn": COLORS["warn"], "danger": COLORS["danger"]}
        fg = colors.get(kind, COLORS["text1"])
        # Show in status bar temporarily
        self.sb_file.config(text=msg, fg=fg)
        self.root.after(3000, lambda: self.sb_file.config(
            fg=COLORS["text2"],
            text=(self.tree_state["docs"].get(self.active_tab,{}).get("name","") or "")))

    # ── KEY BINDINGS ─────────────────────────────────────────────────────────
    def _bind_keys(self):
        r = self.root
        r.bind("<Control-n>",             lambda e: self._new_file())
        r.bind("<Control-o>",             lambda e: self._open_file())
        r.bind("<Control-s>",             lambda e: self._save_file())
        r.bind("<Control-Shift-s>",       lambda e: self._save_file_as())
        r.bind("<Control-d>",             lambda e: (self._clone_lines(), "break"))
        r.bind("<Control-Shift-w>",       lambda e: (self._wrap_wildcard(), "break"))
        r.bind("<Control-h>",             lambda e: self._open_find_with_selection())
        r.bind("<Control-H>",             lambda e: self._open_find_with_selection())
        r.bind("<Control-f>",             lambda e: self._open_find_with_selection())
        r.bind("<F3>",                    lambda e: self._find_next())
        r.bind("<Shift-F3>",              lambda e: self._find_prev())
        r.bind("<F7>",                    lambda e: self._toggle_spell())
        r.bind("<Control-Shift-z>",       lambda e: self._toggle_word_wrap())
        r.bind("<Control-Shift-n>",       lambda e: self._new_folder_dialog())
        r.bind("<Control-comma>",         lambda e: self._show_settings())
        r.bind("<Control-slash>",         lambda e: self._show_hotkeys())
        r.bind("<Alt-Left>",              lambda e: self._nav_back())
        r.bind("<Alt-Right>",             lambda e: self._nav_forward())
        r.bind("<Control-F4>",            lambda e: self._close_tab(self.active_tab) if self.active_tab else None)
        r.bind("<Control-w>",             lambda e: self._close_tab(self.active_tab) if self.active_tab else None)
        r.bind("<Control-W>",             lambda e: self._close_tab(self.active_tab) if self.active_tab else None)
        r.bind("<F2>",                    lambda e: self._rename_via_ctx() if self.ctx_target else None)

    # ── HELPERS ──────────────────────────────────────────────────────────────
    def _simple_input(self, title, prompt, default=""):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.configure(bg=COLORS["bg2"])
        win.resizable(False, False)
        x = self.root.winfo_x() + 300
        y = self.root.winfo_y() + 200
        win.geometry(f"320x120+{x}+{y}")
        win.update_idletasks()
        win.grab_set()
        tk.Label(win, text=prompt, bg=COLORS["bg2"], fg=COLORS["text1"],
                 font=("Segoe UI",10)).pack(pady=(16,4), padx=16, anchor="w")
        var = tk.StringVar(value=default)
        entry = tk.Entry(win, textvariable=var, bg=COLORS["bg3"], fg=COLORS["text0"],
                         insertbackground=COLORS["accent"],
                         relief="flat", bd=1, font=("Segoe UI",11),
                         highlightthickness=1,
                         highlightcolor=COLORS["accent"],
                         highlightbackground=COLORS["border2"])
        entry.pack(fill="x", padx=16)
        entry.select_range(0, "end")
        entry.focus_set()
        result = [None]
        def ok(e=None):
            result[0] = var.get().strip()
            win.destroy()
        def cancel(e=None):
            win.destroy()
        entry.bind("<Return>", ok)
        entry.bind("<Escape>", cancel)
        btn_f = tk.Frame(win, bg=COLORS["bg2"])
        btn_f.pack(pady=8)
        tk.Button(btn_f, text="Cancel", bg=COLORS["bg3"], fg=COLORS["text1"],
                  relief="flat", padx=8, command=cancel).pack(side="left", padx=4)
        tk.Button(btn_f, text="OK", bg=COLORS["accent"], fg=COLORS["bg0"],
                  relief="flat", padx=12, font=("Segoe UI",10,"bold"),
                  command=ok).pack(side="left", padx=4)
        win.wait_window()
        return result[0]

    def _save_session(self):
        self.cfg["open_tabs"] = self.tabs
        self.cfg["active_tab"] = self.active_tab
        save_config(self.cfg)

    def _on_close(self):
        # Flush any pending snapshot so the last few keystrokes are captured
        self._flush_snap_timer()
        unsaved = [self.tree_state["docs"][did]
                   for did in self.tabs
                   if did in self.tree_state["docs"]
                   and self.tree_state["docs"][did].get("modified_unsaved")]
        if unsaved:
            names = ", ".join(d["name"] for d in unsaved[:3])
            if len(unsaved) > 3: names += f" (+{len(unsaved)-3} more)"
            ans = messagebox.askyesnocancel("Unsaved Changes",
                f"Unsaved changes in: {names}\n\nSave all before quitting?")
            if ans is None: return
            if ans:
                for d in unsaved:
                    if d.get("path"):
                        try:
                            with open(d["path"], "w", encoding="utf-8") as f:
                                f.write(d.get("content",""))
                        except Exception:
                            pass
        self.cfg["window_geometry"] = self.root.geometry()
        self.cfg["sidebar_width"]   = self.sidebar.winfo_width()
        self._save_session()
        save_config(self.cfg)
        save_tree_state(self.tree_state)
        self.root.destroy()


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import traceback
    try:
        root = tk.Tk()
        root.title(APP_NAME)

        # Set a custom icon — a stylized "W" card symbol using XBM bitmap (no file needed)
        # XBM is built-in to tkinter, no PIL required
        ICON_XBM = """
#define icon_width 32
#define icon_height 32
static unsigned char icon_bits[] = {
   0xff, 0xff, 0xff, 0xff,
   0xff, 0xff, 0xff, 0xff,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x3c, 0x3c, 0xc0,
   0x03, 0x3c, 0x3c, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0xc3, 0xc3, 0xc0,
   0x03, 0xe7, 0xe7, 0xc0,
   0x03, 0x7e, 0x7e, 0xc0,
   0x03, 0x3c, 0x3c, 0xc0,
   0x03, 0x18, 0x18, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0xff, 0xff, 0xc0,
   0x03, 0xff, 0xff, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x3c, 0x3c, 0xc0,
   0x03, 0x3c, 0x3c, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0x03, 0x00, 0x00, 0xc0,
   0xff, 0xff, 0xff, 0xff,
   0xff, 0xff, 0xff, 0xff};
"""
        # Try PNG icon via PhotoImage (base64 encoded 32x32 purple card icon)
        ICON_B64 = (
            "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1B"
            "AACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAHpSURBVFhH7ZaxTsMwEIb/pqhDGTqx"
            "MiAGJMTABEJiY0JiZ2FkYWViY2JkgIGFgYGBgYWBgYGFgYGBgYGBgYGBgYGB"
            "gYGBgYGBgYGBgYGB"
        )

        # Use a simple drawn icon with Canvas rendered to PhotoImage
        # Create 32x32 icon using tk drawing — dark card with W
        try:
            icon_img = tk.PhotoImage(width=32, height=32)
            # Fill dark background
            for y in range(32):
                for x in range(32):
                    # Border
                    if x < 2 or x > 29 or y < 2 or y > 29:
                        icon_img.put("#7eb8f7", (x, y))
                    # Card background
                    else:
                        icon_img.put("#1a1e28", (x, y))
            # Draw W shape (simple pixel art)
            w_pixels = [
                (5,8),(5,9),(5,10),(5,11),(5,12),(5,13),(5,14),(5,15),(5,16),
                (26,8),(26,9),(26,10),(26,11),(26,12),(26,13),(26,14),(26,15),(26,16),
                (8,16),(9,17),(10,18),(11,19),(12,20),(13,19),(14,18),(15,17),
                (16,16),(17,17),(18,18),(19,19),(20,20),(21,19),(22,18),(23,17),(24,16),
                (6,8),(7,8),(8,8),(23,8),(24,8),(25,8),
            ]
            for px, py in w_pixels:
                if 0 <= px < 32 and 0 <= py < 32:
                    icon_img.put("#a78bfa", (px, py))
            root.iconphoto(True, icon_img)
        except Exception:
            try:
                root.iconbitmap(default="")
            except Exception:
                pass

        app = WildcardEditor(root)
        root.mainloop()
    except Exception:
        log_path = Path(__file__).parent / "wildcard_editor_error.txt"
        with open(log_path, "w") as f:
            traceback.print_exc(file=f)
        try:
            import tkinter.messagebox as mb
            mb.showerror("Wildcard Editor — Startup Error",
                         f"A startup error occurred.\n\nDetails saved to:\n{log_path}\n\n"
                         + traceback.format_exc()[:800])
        except Exception:
            pass
        raise
