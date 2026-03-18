"""
scintilla_text.py
=================
A tk.Text drop-in replacement backed by Scintilla (SciLexer.dll).
Implements exactly the subset of tk.Text API used by wildcard_editor.py.

Usage:
    from scintilla_text import ScintillaText
    # Replace: ed = tk.Text(parent, ...)
    # With:    ed = ScintillaText(parent, ...)

Requirements:
    - SciLexer.dll in the same folder as this file, OR at SCILEXER_PATH below
    - Windows x64, Python 3.8+
    - pip install pywin32  (for win32gui/win32con)
"""

import ctypes
import ctypes.wintypes as wt
import re
import tkinter as tk
from typing import Optional

# ── Path to SciLexer.dll ─────────────────────────────────────────────────────
# Notepad++ ships it; copy SciLexer.dll next to this file or set path here.
import os
_HERE = os.path.dirname(os.path.abspath(__file__))
SCILEXER_CANDIDATES = [
    os.path.join(_HERE, "SciLexer.dll"),
    r"C:\Program Files\Notepad++\SciLexer.dll",
    r"C:\Program Files (x86)\Notepad++\SciLexer.dll",
]

# ── Win32 / Scintilla constants ───────────────────────────────────────────────
WS_CHILD        = 0x40000000
WS_VISIBLE      = 0x10000000
WS_VSCROLL      = 0x00200000
WS_HSCROLL      = 0x00100000
WS_CLIPCHILDREN = 0x02000000
WS_CLIPSIBLINGS = 0x04000000

# Scintilla messages (subset we use)
SCI_GETTEXT             = 2182
SCI_SETTEXT             = 2181
SCI_GETLENGTH           = 2006
SCI_INSERTTEXT          = 2003
SCI_DELETERANGE         = 2645
SCI_CLEARALL            = 2004
SCI_GOTOPOS             = 2025
SCI_ENSUREVISIBLE       = 2232
SCI_SETFIRSTVISIBLELINE = 2613
SCI_LINESONSCREEN       = 2370
SCI_GOTOLINE            = 2024
SCI_SETSEL              = 2160
SCI_GETSELECTIONSTART   = 2143
SCI_GETSELECTIONEND     = 2145
SCI_SETSELECTIONSTART   = 2144
SCI_SETSELECTIONEND     = 2145
SCI_GETANCHOR           = 2009
SCI_GETCURRENTPOS       = 2008
SCI_POSITIONFROMLINE    = 2167
SCI_LINEFROMPOSITION    = 2166
SCI_GETCOLUMN           = 2129
SCI_POSITIONRELATIVE    = 2670
SCI_COUNTCHARACTERS     = 2633
SCI_GETLINECOUNT        = 2154
SCI_LINELENGTH          = 2350
SCI_GETLINE             = 2153
SCI_SCROLLCARET         = 2169
SCI_SETWRAPMODE         = 2268
SCI_GETWRAPMODE         = 2269
SCI_SETLEXER            = 4001
SCI_SETSTYLING          = 2033
SCI_STYLESETFORE        = 2051
SCI_STYLESETBACK        = 2052
SCI_STYLESETFONT        = 2056
SCI_STYLESETSIZE        = 2055
SCI_STARTSTYLING        = 2032
SCI_SETSTYLING          = 2033
SCI_STYLESETBOLD        = 2053
SCI_STYLECLEARALL       = 2050
SCI_SETMARGINWIDTHN     = 2242
SCI_SETMARGINTYPEN      = 2240
SCI_MARGINSETTEXT       = 4004  # not used but reserved
SCI_SETCARETFORE        = 2069
SCI_SETCARETWIDTH       = 2188
SCI_SETSELBACK          = 2068
SCI_SETSELFORE          = 2067
SCI_USEPOPUP            = 2371
SCI_SETCODEPAGE         = 2037
SCI_GETCODEPAGE         = 2137
SCI_BEGINUNDOACTION     = 2078
SCI_ENDUNDOACTION       = 2079
SCI_UNDO                = 2176
SCI_REDO                = 2177
SCI_CANUNDO             = 2174
SCI_CANREDO             = 2016
SCI_EMPTYUNDOBUFFER     = 2175
SCI_ADDUNDOACTION       = 2560
SCI_GETMODIFY           = 2159
SCI_SETSAVEPOINT        = 2014
SCI_FINDTEXT            = 2150
SCI_SEARCHANCHOR        = 2366
SCI_SEARCHNEXT          = 2367
SCI_GETSTYLEAT          = 2010
SCI_SETINDICATORCURRENT = 2500
SCI_INDICSETSTYLE       = 2080
SCI_INDICSETFORE        = 2082
SCI_INDICSETALPHA       = 2523
SCI_INDICSETVALUE       = 2077
SCI_INDICATORFILLRANGE  = 2504
SCI_INDICATORCLEARRANGE = 2505
SCI_INDICSETUNDER       = 2510
SCI_SETREADONLY         = 2171
SCI_ZOOM                = 2373
SCI_CLEARCMDKEY         = 2384
SCI_ASSIGNCMDKEY        = 2380
SCI_BRACEHIGHLIGHT      = 2351
SCI_MARKERDEFINE        = 2040
SCI_MARKERSETBACK       = 2042
SCI_MARKERADD           = 2043
SCI_MARKERDELETE        = 2044
SCI_MARKERDELETEALL     = 2045
SC_MARK_BACKGROUND      = 22
MARKER_PAREN_SPAN       = 1    # line background markers for unmatched paren spans (depths 0-4)
MARKER_PAREN_D          = [1, 2, 3, 4, 5]  # marker number = depth+1
SCI_BRACEBADLIGHT       = 2352
SCI_BRACEMATCH          = 2353
STYLE_BRACELIGHT        = 34
STYLE_BRACEBAD          = 35
SCI_GETFIRSTVISIBLELINE = 2152
SCI_LINESONSCREEN       = 2370
SCWS_INVISIBLE          = 0
SC_WRAP_NONE            = 0
SC_WRAP_WORD            = 1
SC_CP_UTF8              = 65001

# Indicator slots for our highlight types
INDIC_WILDCARD  = 30   # above brackets (9-29)
INDIC_ANGLE     = 31   # above wildcard
INDIC_PAREN     = 10   # unused directly
INDIC_BRACKET   = 11
INDIC_FIND_HL   = 32
INDIC_FIND_CUR  = 33
INDIC_SPELL     = 34
INDIC_WC_ACTIVE   = 35
INDIC_PAREN_SPAN  = 8    # unmatched paren span highlight — below all brackets
INDIC_BRACE_BOLD  = 26   # cursor bracket bold indicator

# Style numbers
STYLE_DEFAULT   = 32

# RECT struct for screen coordinate calculations
class _RECT(ctypes.Structure):
    _fields_ = [("left",ctypes.c_long),("top",ctypes.c_long),
                ("right",ctypes.c_long),("bottom",ctypes.c_long)]

# FindText struct
class _CharRange(ctypes.Structure):
    _fields_ = [("cpMin", ctypes.c_long), ("cpMax", ctypes.c_long)]
class _TextToFind(ctypes.Structure):
    _fields_ = [("chrg", _CharRange), ("lpstrText", ctypes.c_char_p),
                ("chrgText", _CharRange)]

FIND_MATCHCASE  = 4
FIND_WHOLEWORD  = 2
FIND_REGEXP     = (1 << 23)

# ─────────────────────────────────────────────────────────────────────────────

_user32  = ctypes.windll.user32
_kernel32= ctypes.windll.kernel32


_user32.SendMessageW.restype  = ctypes.c_ssize_t
_user32.SendMessageW.argtypes = [wt.HWND, wt.UINT,
                                  ctypes.c_size_t, ctypes.c_ssize_t]

_user32.GetAsyncKeyState.restype  = ctypes.c_short
_user32.GetAsyncKeyState.argtypes = [ctypes.c_int]

_kernel32.GetCurrentThreadId.restype  = wt.DWORD
_kernel32.GetCurrentThreadId.argtypes = []

# Virtual key codes used for hotkey polling
_VK_CONTROL = 0x11
_VK_SHIFT   = 0x10
_VK_S       = 0x53
_VK_W       = 0x57
_VK_F       = 0x46
_VK_H       = 0x48
_VK_N       = 0x4E
_VK_RBUTTON = 0x02
_VK_END     = 0x23

# ── SetWindowSubclass: intercept WM_CHAR control chars at the window-proc level ──
#
# When Ctrl+S is pressed, Windows posts WM_CHAR(0x13) to the Scintilla HWND.
# Scintilla's WM_CHAR handler inserts the character verbatim — that's the DC3.
# SCI_CLEARCMDKEY does NOT help (Ctrl+S was never in Scintilla's command map).
# WH_GETMESSAGE only intercepts queued messages; WM_CHAR sent via SendMessage
# bypasses it. SetWindowSubclass intercepts ALL messages at the window proc —
# posted, sent, and everything in between — making it the definitive fix.
#
# The subclass proc runs in the same thread that owns the window (the main/Tk
# thread), so there are no GIL issues with this ctypes callback approach.

WM_CHAR = 0x0102

_comctl32 = ctypes.windll.comctl32
_comctl32.SetWindowSubclass.restype  = wt.BOOL
_comctl32.SetWindowSubclass.argtypes = [
    wt.HWND, ctypes.c_void_p, ctypes.c_size_t, ctypes.c_size_t]
_comctl32.DefSubclassProc.restype  = ctypes.c_ssize_t
_comctl32.DefSubclassProc.argtypes = [
    wt.HWND, wt.UINT, ctypes.c_size_t, ctypes.c_ssize_t]
_comctl32.RemoveWindowSubclass.restype  = wt.BOOL
_comctl32.RemoveWindowSubclass.argtypes = [
    wt.HWND, ctypes.c_void_p, ctypes.c_size_t]

_SUBCLASSPROC = ctypes.WINFUNCTYPE(
    ctypes.c_ssize_t,
    wt.HWND, wt.UINT, ctypes.c_size_t, ctypes.c_ssize_t,
    ctypes.c_size_t, ctypes.c_size_t,
)

# Keep subclass proc objects alive (one per HWND) so GC doesn't collect them
_subclass_procs: dict = {}   # hwnd -> SUBCLASSPROC instance
_SUBCLASS_ID      = 0x5343   # arbitrary unique ID ("SC") — on Scintilla child

def _make_subclass_proc(hwnd):
    """Return a subclass proc for the given Scintilla HWND that swallows
    WM_CHAR control characters (wParam < 0x20, e.g. 0x13 from Ctrl+S)."""
    def _proc(hwnd_, msg, wp, lp, uid, data):
        if msg == WM_CHAR and wp < 0x20:
            return 0   # swallow — do NOT pass to Scintilla's WM_CHAR handler
        return _comctl32.DefSubclassProc(hwnd_, msg, wp, lp)
    return _SUBCLASSPROC(_proc)

def _register_sci_hwnd(hwnd, sci_widget):
    """Install subclass proc on the Scintilla HWND to block control-char insertion."""
    if hwnd not in _subclass_procs:
        proc = _make_subclass_proc(hwnd)
        _subclass_procs[hwnd] = proc   # keep alive
        _comctl32.SetWindowSubclass(hwnd, proc, _SUBCLASS_ID, 0)

def _unregister_sci_hwnd(hwnd, sci_widget=None):
    """Remove the subclass proc from the Scintilla HWND."""
    proc = _subclass_procs.pop(hwnd, None)
    if proc:
        _comctl32.RemoveWindowSubclass(hwnd, proc, _SUBCLASS_ID)


def _load_scilexer():
    for path in SCILEXER_CANDIDATES:
        if os.path.exists(path):
            try:
                ctypes.cdll.LoadLibrary(path)
                return path
            except Exception:
                pass
    raise RuntimeError(
        "SciLexer.dll not found. Copy it from your Notepad++ folder "
        f"to: {SCILEXER_CANDIDATES[0]}")

_scilexer_loaded = False

def _ensure_loaded():
    global _scilexer_loaded
    if not _scilexer_loaded:
        _load_scilexer()
        _scilexer_loaded = True

def _ptr(buf):
    return ctypes.cast(buf, ctypes.c_void_p).value

# ─────────────────────────────────────────────────────────────────────────────
# ScintillaText
# ─────────────────────────────────────────────────────────────────────────────

class ScintillaText(tk.Frame):
    """
    Drop-in replacement for tk.Text backed by Scintilla.
    Accepts the same constructor kwargs as tk.Text and exposes the
    same method names used by wildcard_editor.py.
    """

    def __init__(self, parent, **kw):
        _ensure_loaded()

        # Extract tk.Text-style kwargs we care about
        self._bg       = kw.pop("bg", kw.pop("background", "#0d0f14"))
        self._fg       = kw.pop("fg", kw.pop("foreground", "#e8eaf0"))
        self._font_fam = "Consolas"
        self._font_sz  = 13
        font = kw.pop("font", None)
        if isinstance(font, tuple) and len(font) >= 2:
            self._font_fam = font[0]
            self._font_sz  = int(font[1])
        self._wrap_mode = SC_WRAP_WORD if kw.pop("wrap","none") == "word" else SC_WRAP_NONE
        self._ins_bg   = kw.pop("insertbackground", "#7eb8f7")
        self._sel_bg   = kw.pop("selectbackground", "#1e3a5f")
        self._sel_fg   = kw.pop("selectforeground", "#e8eaf0")
        # Discard unsupported kwargs
        for k in ["relief","bd","undo","autoseparators","maxundo",
                  "padx","pady","highlightthickness","tabs",
                  "highlightbackground","highlightcolor"]:
            kw.pop(k, None)

        # Keep remaining for Frame
        super().__init__(parent, bg=self._bg, **kw)

        self._hwnd: Optional[int] = None
        self._tag_cfg: dict = {}      # tag_name -> {fg, bg, underline, ...}
        self._tag_ranges: dict = {}   # tag_name -> [(start_bytes, end_bytes)]
        self._modified_flag = False
        self._modified_cb   = None
        self._virtual_cbs: dict = {}
        self._yscroll_cmd   = None
        self._xscroll_cmd   = None
        self._pending_text  = None
        self._dbl_click_cbs = []  # direct callbacks for double-click (Tk can't generate Double events)
        self._hotkey_cbs: dict = {}  # (ctrl, shift, vk) -> [callbacks] for Win32 key polling
        self._find_hl_ranges: list = []  # [(byte_start, byte_end)] of active find_hl/find_cur ranges
        self._TAG_STYLE: dict = {}  # tag -> Scintilla style number for text coloring
        self._styled_ranges: list = []  # [(start, end, style)] for restyling   # content queued before window exists

        # Create the actual Scintilla window after Tk assigns us a real HWND
        self.bind("<Map>", self._on_map, "+")
        self.bind("<Configure>", self._on_configure, "+")

    # ── Window creation ───────────────────────────────────────────────────────

    def _create_mouse_overlay(self):
        """Poll Scintilla cursor position to fire synthetic Tk events.
        This avoids Win32 subclassing (GIL crash) and opaque overlays (blocks input)."""
        self._last_sci_pos = -1
        self._poll_sci_events()

    def _poll_sci_events(self):
        """Detect Scintilla cursor/selection changes and fire matching Tk events."""
        if self._hwnd:
            pos   = self._sci(SCI_GETCURRENTPOS)
            sel_s = self._sci(SCI_GETSELECTIONSTART)
            sel_e = self._sci(SCI_GETSELECTIONEND)
            sel_len = sel_e - sel_s

            if pos != self._last_sci_pos:
                self._last_sci_pos = pos
                try:
                    self.event_generate("<ButtonRelease-1>", x=0, y=0,
                                       rootx=self.winfo_rootx(),
                                       rooty=self.winfo_rooty())
                    self.event_generate("<KeyRelease>", x=0, y=0)
                except Exception:
                    pass
                # Call bracket highlight directly — event_generate may not
                # reliably trigger frame bindings while HWND has Win32 focus
                try:
                    self.update_brace_highlight(pos)
                except Exception:
                    pass

            # Detect double-click: Scintilla selects a word on double-click.
            # When selection appears, move cursor to sel_start so that
            # _on_editor_dbl_click finds the right column, then fire the event.
            last_sel = getattr(self, '_last_sel_len', 0)
            if sel_len > 0 and last_sel == 0 and sel_len < 200:
                sel_text = self._get_bytes(sel_s, sel_e).decode("utf-8", errors="replace")
                # Only fire on real words (not wrap chars like ~~, not spaces)
                if sel_text and len(sel_text) > 1 and " " not in sel_text and "\n" not in sel_text:
                    mid = sel_s + (sel_e - sel_s) // 2
                    self._sci(SCI_GOTOPOS, mid)
                    # event_generate("<Double-Button-1>") is illegal in Tk.
                    # Call registered dbl-click callbacks directly.
                    for cb in self._dbl_click_cbs:
                        try:
                            cb(None)
                        except Exception:
                            pass
            # Track selection for <<SelectionCleared>> event only
            _prev_sel_s = getattr(self, '_prev_sel_s', -1)
            _prev_sel_e = getattr(self, '_prev_sel_e', -1)
            if sel_s != _prev_sel_s or sel_e != _prev_sel_e:
                if sel_e <= sel_s and _prev_sel_e > _prev_sel_s:
                    self._fire_virtual("<<SelectionCleared>>")
                self._prev_sel_s = sel_s
                self._prev_sel_e = sel_e
            self._last_sel_len = sel_len

            # ── Hotkey polling ────────────────────────────────────────────────
            # Tk bindings on the ScintillaText frame never fire while the Win32
            # HWND has focus.  We detect hotkeys here via GetAsyncKeyState and
            # call registered callbacks directly — same pattern as dbl-click.
            if self._hotkey_cbs and _user32.GetFocus() == self._hwnd:
                ctrl  = bool(_user32.GetAsyncKeyState(_VK_CONTROL) & 0x8000)
                shift = bool(_user32.GetAsyncKeyState(_VK_SHIFT)   & 0x8000)
                s_dn  = bool(_user32.GetAsyncKeyState(_VK_S)       & 0x8000)
                w_dn  = bool(_user32.GetAsyncKeyState(_VK_W)       & 0x8000)
                f_dn  = bool(_user32.GetAsyncKeyState(_VK_F)       & 0x8000)
                h_dn  = bool(_user32.GetAsyncKeyState(_VK_H)       & 0x8000)
                n_dn  = bool(_user32.GetAsyncKeyState(_VK_N)       & 0x8000)
                _prev_s = getattr(self, '_prev_s_down', False)
                _prev_w = getattr(self, '_prev_w_down', False)
                _prev_f = getattr(self, '_prev_f_down', False)
                _prev_h = getattr(self, '_prev_h_down', False)
                _prev_n = getattr(self, '_prev_n_down', False)
                # Fire on the leading edge (key goes down), not while held
                if ctrl and s_dn and not _prev_s:
                    key = (True, shift, _VK_S)
                    for cb in self._hotkey_cbs.get(key, []):
                        try:
                            cb(None)
                        except Exception:
                            pass
                if ctrl and w_dn and not _prev_w:
                    key = (True, shift, _VK_W)
                    for cb in self._hotkey_cbs.get(key, []):
                        try:
                            cb(None)
                        except Exception:
                            pass
                if ctrl and f_dn and not _prev_f:
                    key = (True, shift, _VK_F)
                    for cb in self._hotkey_cbs.get(key, []):
                        try:
                            cb(None)
                        except Exception:
                            pass
                if ctrl and h_dn and not _prev_h:
                    key = (True, shift, _VK_H)
                    for cb in self._hotkey_cbs.get(key, []):
                        try:
                            cb(None)
                        except Exception:
                            pass
                if ctrl and n_dn and not _prev_n and not shift:
                    key = (True, False, _VK_N)
                    for cb in self._hotkey_cbs.get(key, []):
                        try:
                            cb(None)
                        except Exception:
                            pass
                self._prev_s_down = s_dn
                self._prev_w_down = w_dn
                self._prev_f_down = f_dn
                self._prev_h_down = h_dn
                self._prev_n_down = n_dn

            # ── Right-click polling ───────────────────────────────────────────
            # SCI_USEPOPUP(0) disables Scintilla's own context menu. We detect
            # right-click here (main thread, GIL-safe) and fire <Button-3> on
            # the Tk frame so _editor_right_click fires normally.
            rb_dn = bool(_user32.GetAsyncKeyState(_VK_RBUTTON) & 0x8000)
            _prev_rb = getattr(self, '_prev_rb_down', False)
            if not rb_dn and _prev_rb:
                try:
                    class _PT(ctypes.Structure):
                        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
                    pt = _PT()
                    _user32.GetCursorPos(ctypes.byref(pt))
                    # Only fire if cursor is within our widget bounds at release time
                    _user32.ScreenToClient(self._hwnd, ctypes.byref(pt))
                    w = self.winfo_width()
                    h = self.winfo_height()
                    if (0 <= pt.x <= w and 0 <= pt.y <= h
                            and _user32.GetFocus() == self._hwnd):
                        self.event_generate("<Button-3>", x=pt.x, y=pt.y,
                                           rootx=self.winfo_rootx() + pt.x,
                                           rooty=self.winfo_rooty() + pt.y)
                except Exception:
                    pass
            self._prev_rb_down = rb_dn

            # ── End key polling ───────────────────────────────────────────────
            # SCI_CLEARCMDKEY removed Scintilla's default End (go to line end).
            # Poll here and fire <<EndKey>> so wildcard_editor jumps to doc end.
            end_dn = bool(_user32.GetAsyncKeyState(_VK_END) & 0x8000)
            _prev_end = getattr(self, '_prev_end_down', False)
            if end_dn and not _prev_end:
                try:
                    self._fire_virtual("<<EndKey>>")
                except Exception:
                    pass
            self._prev_end_down = end_dn


            # Tk yscrollcommand callback. Poll every 16ms and call it when the
            # scroll position changes so the line-number canvas stays in sync.
            if self._yscroll_cmd:
                first_vis  = self._sci(SCI_GETFIRSTVISIBLELINE)
                total      = max(self._sci(SCI_GETLINECOUNT), 1)
                visible    = max(self._sci(SCI_LINESONSCREEN), 1)
                # Fire line number redraw if line count changed
                if total != getattr(self, '_last_line_count', -1):
                    self._last_line_count = total
                    self._fire_virtual("<<LineCountChanged>>")
                first_frac = first_vis / total
                last_frac  = min((first_vis + visible) / total, 1.0)
                if (first_frac, last_frac) != getattr(self, '_last_scroll_frac', None):
                    self._last_scroll_frac = (first_frac, last_frac)
                    try:
                        self._yscroll_cmd(first_frac, last_frac)
                    except Exception:
                        pass

        self.after(32, self._poll_sci_events)



    def _on_map(self, event=None):
        if not self._hwnd:
            self.update_idletasks()
            parent_hwnd = self.winfo_id()
            w = max(self.winfo_width(), 100)
            h = max(self.winfo_height(), 100)
            self._hwnd = _user32.CreateWindowExW(
                0, "Scintilla", "",
                WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                0, 0, w, h,
                parent_hwnd, 0,
                _kernel32.GetModuleHandleW(None), 0)
            if not self._hwnd:
                err = _kernel32.GetLastError()
                raise RuntimeError(f"Failed to create Scintilla window (err={err})")
            self._setup()
            self._flush_pending()
            self._create_mouse_overlay()
            # Re-trigger highlights — they ran before HWND existed at startup
            self.after(50, lambda: self._fire_virtual("<<ScintillaReady>>"))
        else:
            # HWND already exists — tab is being re-shown after pack_forget.
            self._flush_pending()
            self.after(30, lambda: self._fire_virtual("<<ScintillaReady>>"))
        self.after(50,  self._on_configure)
        self.after(150, self._on_configure)
        self.after(500, self._on_configure)

    def _flush_pending(self):
        """Load queued content into Scintilla. No-op if nothing is pending."""
        if not self._pending_text:
            return
        text = self._pending_text
        self._pending_text = None
        # Normalise line endings — Scintilla is set to SC_EOL_LF and get() strips
        # \r, so the buffer must contain only \n to keep byte positions consistent.
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        raw = text.encode("utf-8")
        buf = ctypes.create_string_buffer(raw + b'\x00')
        self._sci(SCI_SETTEXT, 0, _ptr(buf))
        self._sci(SCI_SETSAVEPOINT)
        self._sci(SCI_EMPTYUNDOBUFFER)
        length = self._sci(SCI_GETLENGTH)

    def _on_configure(self, event=None):
        if not self._hwnd:
            return
        self.update_idletasks()
        w = max(self.winfo_width(),  10)
        h = max(self.winfo_height(), 10)
        _user32.SetWindowPos(self._hwnd, 0, 0, 0, w, h, 0x0004 | 0x0010)

    def _sci(self, msg, wp=0, lp=0):
        return _user32.SendMessageW(self._hwnd, msg, wp, lp)

    def _setup(self):
        sci = self._sci
        sci(SCI_SETCODEPAGE, SC_CP_UTF8)
        sci(SCI_SETWRAPMODE, self._wrap_mode)
        sci(SCI_USEPOPUP, 0)   # disable Scintilla's own menu; right-click handled via poll
        # Allow scrolling past the last line (like scroll padding in tk.Text mode)
        sci(2277, 0)   # SCI_SETENDATLASTLINE(0) — don't clamp scroll at last line
        sci(SCI_SETCARETWIDTH, 2)
        sci(2031, 2)   # SCI_SETEOLMODE SC_EOL_LF=2 — force Unix line endings

        # Clear all margins (line numbers etc — we draw our own)
        for i in range(5):
            sci(SCI_SETMARGINWIDTHN, i, 0)

        # Apply theme to style 32 (template) AND style 0 (used for all unstyled text).
        # We cannot use SCI_STYLECLEARALL to propagate 32->0 because it wipes document
        # content in the Notepad++ 7.x bundled Scintilla. Set both explicitly instead.
        self._apply_style(STYLE_DEFAULT, self._fg, self._bg, self._font_fam, self._font_sz)
        self._apply_style(0,             self._fg, self._bg, self._font_fam, self._font_sz)
        sci(SCI_SETCARETFORE, self._color_to_bgr(self._ins_bg))
        sci(SCI_SETSELFORE, 1, self._color_to_bgr(self._sel_fg))
        sci(SCI_SETSELBACK, 1, self._color_to_bgr(self._sel_bg))

        # Set up indicators
        self._setup_indicators()

        # Suppress Scintilla's built-in Ctrl+S (inserts DC3).
        # Key code packed as: (keycode & 0xFFFF) | (modifiers << 16)
        # SCMOD_CTRL = 2. So Ctrl+S = (0x53) | (2 << 16)
        _ctrl_s_cmd = (ord('S')) | (2 << 16)
        sci(SCI_CLEARCMDKEY, _ctrl_s_cmd)
        # Clear End key so we can override it to jump to document end
        sci(SCI_CLEARCMDKEY, 0x23)   # VK_END, no modifiers

        # Subclass the Scintilla HWND to intercept WM_CHAR control characters
        # (e.g. 0x13 from Ctrl+S) before Scintilla's WM_CHAR handler inserts them.
        # SetWindowSubclass runs in the main thread — no GIL issues.
        _register_sci_hwnd(self._hwnd, self)

        # Notify parent of content changes via WM_NOTIFY → we poll instead
        # Actually Scintilla sends SCN_MODIFIED; we wire it via a timer
        self._poll_modified()

    def _apply_style(self, style_num, fg, bg, font, size):
        sci = self._sci
        sci(SCI_STYLESETFORE, style_num, self._color_to_bgr(fg))
        sci(SCI_STYLESETBACK, style_num, self._color_to_bgr(bg))
        font_buf = ctypes.create_string_buffer(font.encode("utf-8"))
        sci(SCI_STYLESETFONT, style_num, _ptr(font_buf))
        sci(SCI_STYLESETSIZE, style_num, size)

    def _setup_indicators(self):
        # Style 8 = INDIC_STRAIGHTBOX (filled rect, Scintilla 3.0+)
        # Style 1 = INDIC_SQUIGGLE
        STRAIGHTBOX = 8
        SQUIGGLE_S  = 1

        def _bgr(hex_color):
            return self._color_to_bgr(hex_color)

        def _cfg_bg(tag, default):
            return _bgr(self._tag_cfg.get(tag, {}).get("background", default))

        def _cfg_fg(tag, default):
            return _bgr(self._tag_cfg.get(tag, {}).get("foreground", default))

        # Main indicators — use colors from tag_configure if already called
        # Alpha 255 = fully opaque, but STRAIGHTBOX blends with text bg.
        # We use high alpha so colors are clearly visible on the dark theme.
        for indic, bgr, alpha, style in [
            (INDIC_WILDCARD,   _cfg_bg("wildcard",  "#2a1f42"), 220, STRAIGHTBOX),
            (INDIC_ANGLE,      _cfg_bg("hl_angle",  "#1a6b2a"), 220, STRAIGHTBOX),
            (INDIC_PAREN,      0x1D2734,                        200, STRAIGHTBOX),
            (INDIC_BRACKET,    0xF7A07E,                        200, STRAIGHTBOX),
            (INDIC_FIND_HL,    _cfg_bg("find_hl",   "#1e3a1e"), 240, STRAIGHTBOX),
            (INDIC_FIND_CUR,   _cfg_bg("find_cur",  "#f59e0b"), 255, STRAIGHTBOX),
            (INDIC_SPELL,      _cfg_fg("spell_err", "#f87171"), 255, SQUIGGLE_S),
            (INDIC_WC_ACTIVE,  _cfg_bg("wc_active", "#7eb8f7"), 240, STRAIGHTBOX),
            (INDIC_BRACE_BOLD, 0xffe500,                        220, STRAIGHTBOX),
        ]:
            self._sci(SCI_INDICSETSTYLE, indic, style)
            self._sci(SCI_INDICSETFORE,  indic, bgr)
            self._sci(SCI_INDICSETALPHA, indic, alpha)
            self._sci(SCI_INDICSETUNDER, indic, 1)
        # INDIC_PAREN_SPAN: indicator for character-level coloring (stacks with bracket colors)
        self._sci(SCI_INDICSETSTYLE, INDIC_PAREN_SPAN, STRAIGHTBOX)
        self._sci(SCI_INDICSETFORE,  INDIC_PAREN_SPAN, 0x453424)
        self._sci(SCI_INDICSETALPHA, INDIC_PAREN_SPAN, 180)
        self._sci(SCI_INDICSETUNDER, INDIC_PAREN_SPAN, 1)
        # MARKER_PAREN_D[0-4]: line background markers, one per depth
        self._sci(SCI_MARKERDEFINE, 1, SC_MARK_BACKGROUND)
        self._sci(SCI_MARKERSETBACK, 1, 0x453424)
        self._sci(SCI_MARKERDEFINE, 2, SC_MARK_BACKGROUND)
        self._sci(SCI_MARKERSETBACK, 2, 0x56412b)
        self._sci(SCI_MARKERDEFINE, 3, SC_MARK_BACKGROUND)
        self._sci(SCI_MARKERSETBACK, 3, 0x674e32)
        self._sci(SCI_MARKERDEFINE, 4, SC_MARK_BACKGROUND)
        self._sci(SCI_MARKERSETBACK, 4, 0x785b39)
        self._sci(SCI_MARKERDEFINE, 5, SC_MARK_BACKGROUND)
        self._sci(SCI_MARKERSETBACK, 5, 0x896840)

        # Named styles for text coloring (SCI_SETSTYLING approach, works in all versions)
        # Style 1 = wildcard text:  fg=#c4b5fd on bg=#2a1f42
        # Style 2 = angle text:     fg=#b8e8e8 on bg=#0d2524
        # Style 3 = find_hl text:   fg=#34d399 on bg=#1e3a1e
        # Style 4 = find_cur text:  fg=#0d0f14 on bg=#f59e0b
        # Style 5 = wc_active text: fg=#0d0f14 on bg=#7eb8f7
        # Style 6 = spell_err text: fg=#f87171 on default bg (underline via indicator)
        wc_bg  = _cfg_bg("wildcard",  "#2a1f42")
        wc_fg  = _cfg_fg("wildcard",  "#c4b5fd")
        an_bg  = _cfg_bg("hl_angle",  "#1a6b2a")
        an_fg  = _cfg_fg("hl_angle",  "#b8e8e8")
        fh_bg  = _cfg_bg("find_hl",   "#1e3a1e")
        fh_fg  = _cfg_fg("find_hl",   "#34d399")
        fc_bg  = _cfg_bg("find_cur",  "#f59e0b")
        fc_fg  = _cfg_fg("find_cur",  "#0d0f14")
        wa_bg  = _cfg_bg("wc_active", "#7eb8f7")
        wa_fg  = _cfg_fg("wc_active", "#0d0f14")
        se_fg  = _cfg_fg("spell_err", "#f87171")
        def _sty(n, fg, bg):
            self._sci(SCI_STYLESETFORE, n, fg)
            self._sci(SCI_STYLESETBACK, n, bg)
            self._sci(SCI_STYLESETFONT, n, _ptr(ctypes.create_string_buffer(self._font_fam.encode("utf-8"))))
            self._sci(SCI_STYLESETSIZE, n, self._font_sz)
        _sty(1, wc_fg, wc_bg)
        _sty(2, an_fg, an_bg)
        _sty(3, fh_fg, fh_bg)
        _sty(4, fc_fg, fc_bg)
        _sty(5, wa_fg, wa_bg)
        _sty(6, se_fg, self._color_to_bgr(self._bg))
        # STYLE_BRACELIGHT (34): used by SCI_BRACEHIGHLIGHT — white bold for matched bracket
        self._sci(SCI_STYLESETFORE, STYLE_BRACELIGHT, 0xFFFFFF)
        self._sci(SCI_STYLESETBACK, STYLE_BRACELIGHT, self._color_to_bgr(self._bg))
        self._sci(SCI_STYLESETBOLD, STYLE_BRACELIGHT, 1)
        # STYLE_BRACEBAD (35): unmatched bracket — red
        self._sci(SCI_STYLESETFORE, STYLE_BRACEBAD, 0x6060FF)
        self._sci(SCI_STYLESETBACK, STYLE_BRACEBAD, self._color_to_bgr(self._bg))
        self._sci(SCI_STYLESETBOLD, STYLE_BRACEBAD, 1)
        self._TAG_STYLE = {
            "wildcard": 1, "hl_angle": 2, "find_hl": 3,
            "find_cur": 4, "wc_active": 5, "spell_err": 6,
        }

        # Depth-bracket indicators: paren_d0-4 (9-13), sqbr_d0-4 (14-18), curly_d0-4 (19-23)
        # Colors match wildcard_editor.py _paren_bgs / _sqbr_bgs (converted #RRGGBB → BGR)
        paren_bgs = [0x453424,0x56412b,0x674e32,0x785b39,0x896840,0x9a7547,0x9a7547]
        sqbr_bgs  = [0x162034,0x182743,0x1a2e52,0x1c3561,0x1e3c70,0x20437f,0x20437f]
        curly_bgs = paren_bgs  # same as paren per wildcard_editor.py
        for base, bgs in ((9, paren_bgs), (14, sqbr_bgs), (19, curly_bgs)):
            for i, bgr in enumerate(bgs[:5]):
                self._sci(SCI_INDICSETSTYLE, base+i, STRAIGHTBOX)
                self._sci(SCI_INDICSETFORE,  base+i, bgr)
                self._sci(SCI_INDICSETALPHA, base+i, 255)
                self._sci(SCI_INDICSETUNDER, base+i, 1)

    # ── Polling for modification ───────────────────────────────────────────────
    def _poll_modified(self):
        if self._hwnd:
            mod = bool(self._sci(SCI_GETMODIFY))
            if mod != self._modified_flag:
                self._modified_flag = mod
                self._fire_virtual("<<Modified>>")
        self.after(200, self._poll_modified)

    def _fire_virtual(self, event_name):
        for cb in self._virtual_cbs.get(event_name, []):
            try:
                cb(None)
            except Exception:
                pass

    # ── Color helpers ─────────────────────────────────────────────────────────
    @staticmethod
    def _color_to_bgr(hex_color: str) -> int:
        """Convert #RRGGBB to Scintilla's BGR int."""
        h = hex_color.lstrip("#")
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        r, g, b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        return (b << 16) | (g << 8) | r

    # ── Index conversion ──────────────────────────────────────────────────────
    def _pos_to_index(self, byte_pos: int) -> str:
        """Convert Scintilla byte position → 'line.col' string."""
        if not self._hwnd:
            return "1.0"
        line = self._sci(SCI_LINEFROMPOSITION, byte_pos)
        line_start = self._sci(SCI_POSITIONFROMLINE, line)
        # Column = character count from line start
        line_bytes = self._get_bytes(line_start, byte_pos)
        col = len(line_bytes.decode("utf-8", errors="replace"))
        return f"{line+1}.{col}"

    def _index_to_pos(self, index: str) -> int:
        """Convert tk.Text index string → Scintilla byte position."""
        if not self._hwnd:
            return 0
        index = str(index).strip()

        # Special names
        if index == "insert":
            return self._sci(SCI_GETCURRENTPOS)
        if index in ("end", "end-1c"):
            pos = self._sci(SCI_GETLENGTH)
            if index == "end-1c" and pos > 0:
                pos -= 1
            return pos
        if index == "sel.first":
            return self._sci(SCI_GETSELECTIONSTART)
        if index == "sel.last":
            return self._sci(SCI_GETSELECTIONEND)
        if index == "1.0":
            return 0

        # line.col format
        if re.match(r'^\d+\.\d+$', index):
            line, col = index.split(".")
            line_start = self._sci(SCI_POSITIONFROMLINE, int(line)-1)
            if line_start < 0:
                return self._sci(SCI_GETLENGTH)
            # Advance col characters
            pos = line_start
            for _ in range(int(col)):
                next_pos = self._sci(SCI_POSITIONRELATIVE, pos, 1)
                if next_pos <= pos:
                    break
                pos = next_pos
            return pos

        # "1.0+Nc" offset format — fast path using pre-cached content bytes
        m = re.match(r'^(.+?)\+(\d+)c$', index)
        if m:
            base = self._index_to_pos(m.group(1))
            n    = int(m.group(2))
            # Fast: get all bytes from base, encode n chars to get byte offset
            total = self._sci(SCI_GETLENGTH)
            if total <= 0:
                return base
            remaining = total - base
            if remaining <= 0:
                return total
            chunk = min(remaining, n * 4 + 16)  # UTF-8: max 4 bytes/char
            buf = ctypes.create_string_buffer(chunk + 1)
            class _TR(ctypes.Structure):
                _fields_ = [("cpMin",ctypes.c_long),("cpMax",ctypes.c_long),
                             ("lpstrText",ctypes.c_char_p)]
            tr = _TR(); tr.cpMin = base; tr.cpMax = base + chunk
            tr.lpstrText = ctypes.cast(buf, ctypes.c_char_p)
            _user32.SendMessageW(self._hwnd, 2162, 0, _ptr(ctypes.pointer(tr)))
            raw = buf.raw[:chunk]
            try:
                # Advance exactly n unicode characters through the UTF-8 bytes
                byte_offset = len(raw.decode('utf-8', errors='replace')[:n].encode('utf-8'))
            except Exception:
                byte_offset = min(n, chunk)
            return base + byte_offset

        # "@x,y" pixel coordinate
        m = re.match(r'^@(\d+),(\d+)$', index)
        if m:
            # Scintilla uses SCI_POSITIONFROMPOINTCLOSE
            # msg 2023 = SCI_POSITIONFROMPOINT
            x, y = int(m.group(1)), int(m.group(2))
            return _user32.SendMessageW(self._hwnd, 2023, x, y)

        # "index wordstart" / "index wordend" / "index linestart" / "index lineend"
        for suffix, handler in [
            (" wordstart", lambda p: self._word_start(p)),
            (" wordend",   lambda p: self._word_end(p)),
            (" linestart", lambda p: self._sci(SCI_POSITIONFROMLINE,
                                               self._sci(SCI_LINEFROMPOSITION, p))),
            (" lineend",   lambda p: self._line_end(p)),
        ]:
            if index.endswith(suffix):
                base = self._index_to_pos(index[:-len(suffix)])
                return handler(base)

        return 0

    def _word_start(self, pos):
        content = self._get_all_text()
        char_pos = len(self._get_bytes(0, pos).decode("utf-8", errors="replace"))
        while char_pos > 0 and content[char_pos-1:char_pos].isalnum() or \
              (char_pos > 0 and content[char_pos-1] == '_'):
            char_pos -= 1
        return self._char_to_pos(char_pos)

    def _word_end(self, pos):
        content = self._get_all_text()
        char_pos = len(self._get_bytes(0, pos).decode("utf-8", errors="replace"))
        n = len(content)
        while char_pos < n and (content[char_pos].isalnum() or content[char_pos] == '_'):
            char_pos += 1
        return self._char_to_pos(char_pos)

    def _line_end(self, pos):
        line = self._sci(SCI_LINEFROMPOSITION, pos)
        line_start = self._sci(SCI_POSITIONFROMLINE, line)
        line_len   = self._sci(SCI_LINELENGTH, line)
        end = line_start + line_len
        # Trim trailing \r\n
        all_len = self._sci(SCI_GETLENGTH)
        while end > line_start:
            prev = end - 1
            if prev >= all_len:
                end -= 1
                continue
            buf = ctypes.create_string_buffer(2)
            _user32.SendMessageW(self._hwnd, SCI_GETTEXT, 2,
                                 _ptr(buf))
            # crude — just trim by checking the byte
            b = self._get_bytes(prev, end)
            if b in (b'\r', b'\n'):
                end -= 1
            else:
                break
        return end

    def _char_to_pos(self, char_idx: int) -> int:
        content = self._get_all_text()
        return len(content[:char_idx].encode("utf-8"))

    def _get_bytes(self, start: int, end: int) -> bytes:
        if end <= start:
            return b""
        n = end - start
        buf = ctypes.create_string_buffer(n + 1)
        # Use SCI_GETTEXT on a range via a TextRange struct
        class TextRange(ctypes.Structure):
            _fields_ = [("cpMin", ctypes.c_long), ("cpMax", ctypes.c_long),
                        ("lpstrText", ctypes.c_char_p)]
        tr = TextRange()
        tr.cpMin    = start
        tr.cpMax    = end
        tr.lpstrText= ctypes.cast(buf, ctypes.c_char_p)
        _user32.SendMessageW(self._hwnd, 2162,  # SCI_GETTEXTRANGE
                             0, _ptr(ctypes.pointer(tr)))
        return buf.raw[:n]

    def _ensure_hwnd(self):
        """Create the Scintilla window if not yet done — but only if the frame
        has real geometry. If called pre-pack (frame is 1x1 or unmapped), we
        defer: content queued as _pending_text will be flushed by _on_map."""
        if self._hwnd:
            return
        self.update_idletasks()
        parent_hwnd = self.winfo_id()
        if not parent_hwnd:
            return  # not ready, _on_map will handle it
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 10 or h < 10:
            return  # frame not laid out yet — defer to _on_map
        self._hwnd = _user32.CreateWindowExW(
            0, "Scintilla", "",
            WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
            0, 0, w, h,
            parent_hwnd, 0,
            _kernel32.GetModuleHandleW(None), 0)
        if self._hwnd:
            self._setup()
            self._flush_pending()
            self.after(50,  self._on_configure)
            self.after(150, self._on_configure)
            self.after(400, self._on_configure)

    def _get_all_text(self) -> str:
        self._ensure_hwnd()
        if not self._hwnd:
            return ""
        n = self._sci(SCI_GETLENGTH)
        if n <= 0:
            return ""
        buf = ctypes.create_string_buffer(n + 1)
        self._sci(SCI_GETTEXT, n + 1, _ptr(buf))
        result = buf.raw[:n].decode("utf-8", errors="replace")
        return result.replace("\r\n", "\n").replace("\r", "\n")

    # ── tk.Text API ───────────────────────────────────────────────────────────

    def get(self, index1="1.0", index2=None) -> str:
        if not self._hwnd:
            return ""
        if index2 is None:
            index2 = "end"
        p1 = self._index_to_pos(index1)
        p2 = self._index_to_pos(index2)
        if p2 <= p1:
            return ""
        return self._get_bytes(p1, p2).decode("utf-8", errors="replace").replace("\r\n", "\n").replace("\r", "\n")

    def insert(self, index, text, *tags):
        # _pad is a tk.Text scroll-padding hack — not needed with Scintilla
        if "_pad" in tags:
            return
        self._ensure_hwnd()
        if isinstance(text, bytes):
            text = text.decode("utf-8", errors="replace")
        # Normalise line endings to match SC_EOL_LF mode
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        if not self._hwnd:
            # Pre-pack: HWND not ready. Queue content; _on_map/_flush_pending will load it.
            if index in ("1.0", "end") and not tags and text:
                self._pending_text = (self._pending_text or "") + text
            return
        raw = text.encode("utf-8")
        # Use SCI_SETTEXT for the initial document load (empty doc, no tags).
        # SCI_INSERTTEXT at pos 0 into a fresh document can silently succeed
        # yet not display — SCI_SETTEXT is the reliable path.
        if not tags and self._sci(SCI_GETLENGTH) == 0 and index in ("1.0", "end"):
            buf = ctypes.create_string_buffer(raw + b'\x00')
            self._sci(SCI_SETTEXT, 0, _ptr(buf))
            self._sci(SCI_SETSAVEPOINT)
            self._sci(SCI_EMPTYUNDOBUFFER)
            return
        if index == "end":
            pos = self._sci(SCI_GETLENGTH)
        else:
            pos = self._index_to_pos(index)
        buf = ctypes.create_string_buffer(raw + b'\x00')
        self._sci(SCI_INSERTTEXT, pos, _ptr(buf))
        after_len = self._sci(SCI_GETLENGTH)
        for tag in tags:
            end_pos = pos + len(raw)
            self.tag_add(tag, self._pos_to_index(pos),
                              self._pos_to_index(end_pos))

    def delete(self, index1, index2=None):
        self._ensure_hwnd()
        if not self._hwnd:
            self._pending_text = None
            return
        if index1 == "1.0" and index2 in ("end", "end-1c", None):
            self._sci(SCI_CLEARALL)
            self._pending_text = None
            return
        p1 = self._index_to_pos(index1)
        if index2 is None:
            p2 = self._sci(SCI_POSITIONRELATIVE, p1, 1)
        else:
            p2 = self._index_to_pos(index2)
        if p2 > p1:
            self._sci(SCI_DELETERANGE, p1, p2 - p1)

    def index(self, index: str) -> str:
        """Return canonical 'line.col' for given index."""
        if not self._hwnd:
            return "1.0"
        pos = self._index_to_pos(index)
        result = self._pos_to_index(pos)
        return result

    def mark_set(self, mark, index):
        if not self._hwnd:
            return
        if mark == "insert":
            pos = self._index_to_pos(index)
            self._sci(SCI_GOTOPOS, pos)

    def mark_names(self):
        return ["insert", "current"]

    def see(self, index):
        if not self._hwnd:
            return
        pos = self._index_to_pos(index)
        line = self._sci(SCI_LINEFROMPOSITION, pos)
        self._sci(SCI_ENSUREVISIBLE, line)
        first = self._sci(SCI_GETFIRSTVISIBLELINE)
        visible = self._sci(SCI_LINESONSCREEN)
        last = first + visible - 1
        if line < first or line > last:
            # Center the target line in the viewport
            self._sci(SCI_SETFIRSTVISIBLELINE, max(0, line - visible // 2))
        self._sci(SCI_GOTOPOS, pos)

    def compare(self, index1, op, index2) -> bool:
        p1 = self._index_to_pos(index1)
        p2 = self._index_to_pos(index2)
        return {"<": p1 < p2, "<=": p1 <= p2, "==": p1 == p2,
                ">=": p1 >= p2, ">": p1 > p2, "!=": p1 != p2}.get(op, False)

    def config(self, **kw):
        # yscrollcommand/xscrollcommand don't need the HWND — store them always
        if "yscrollcommand" in kw:
            self._yscroll_cmd = kw.pop("yscrollcommand")
        if "xscrollcommand" in kw:
            self._xscroll_cmd = kw.pop("xscrollcommand")
        if not self._hwnd:
            return
        if "font" in kw:
            f = kw["font"]
            if isinstance(f, tuple) and len(f) >= 2:
                self._font_fam = f[0]; self._font_sz = int(f[1])
                # Apply font to default style only — do NOT call SCI_STYLECLEARALL
                # as it wipes document content
                self._apply_style(STYLE_DEFAULT, self._fg, self._bg,
                                  self._font_fam, self._font_sz)
        if "wrap" in kw:
            mode = SC_WRAP_WORD if kw["wrap"] == "word" else SC_WRAP_NONE
            self._sci(SCI_SETWRAPMODE, mode)
        if "state" in kw:
            self._sci(SCI_SETREADONLY, 1 if kw["state"] == "disabled" else 0)

    configure = config

    def cget(self, key):
        if key in ("wrap",):
            m = self._sci(SCI_GETWRAPMODE)
            return "word" if m == SC_WRAP_WORD else "none"
        return None

    def yview(self, *args):
        pass  # Scintilla handles its own scrollbar

    def yview_scroll(self, number, what):
        """Scroll Scintilla vertically by lines or pages."""
        if not self._hwnd:
            return
        if what == "units":
            # SCI_LINESCROLL(columns, lines)
            _user32.SendMessageW(self._hwnd, 2300, 0, int(number))
        elif what == "pages":
            cmd = 2312 if number > 0 else 2313  # SCI_PAGEDOWN / SCI_PAGEUP
            for _ in range(abs(int(number))):
                _user32.SendMessageW(self._hwnd, cmd, 0, 0)

    def yview_moveto(self, fraction):
        """Scroll to a fractional position (0.0=top, 1.0=bottom)."""
        if not self._hwnd:
            return
        total_lines = self._sci(SCI_GETLINECOUNT)
        line = int(fraction * total_lines)
        _user32.SendMessageW(self._hwnd, 2024, 0, line)  # SCI_GOTOLINE

    def xview(self, *args):
        pass

    def focus_set(self):
        """Focus the Scintilla window."""
        if self._hwnd:
            _user32.SetFocus(self._hwnd)
        super().focus_set()

    def destroy(self):
        """Unregister HWND from the WM_CHAR hook before tearing down."""
        if self._hwnd:
            _unregister_sci_hwnd(self._hwnd, self)
        super().destroy()

    def tag_names(self, index=None):
        """Return tag names active at index (used for spell-check detection)."""
        if not self._hwnd or index is None:
            return ()
        pos = self._index_to_pos(str(index))
        active = []
        # Check spell_err via Python-tracked ranges (SCI_INDICATORVALUEAT is
        # unreliable for this indicator in Notepad++ 7.x Scintilla builds)
        for s, e in self._tag_ranges.get("spell_err", []):
            if s <= pos < e:
                active.append("spell_err")
                break
        # Check other indicators via SCI_INDICATORVALUEAT
        for tag, indic in {**self._TAG_TO_INDIC, **self._TAG_TO_INDIC_DEPTH}.items():
            if tag == "spell_err":
                continue
            val = _user32.SendMessageW(self._hwnd, 2506, indic, pos)
            if val:
                active.append(tag)
        return tuple(active)

    def event_generate(self, sequence, **kw):
        """Generate a Tk event — delegate to Frame."""
        try:
            super().event_generate(sequence, **kw)
        except Exception:
            pass

    # ── Tag API ───────────────────────────────────────────────────────────────

    _TAG_TO_INDIC = {
        "wildcard":  INDIC_WILDCARD,
        "hl_angle":  INDIC_ANGLE,
        "find_hl":   INDIC_FIND_HL,
        "find_cur":  INDIC_FIND_CUR,
        "spell_err": INDIC_SPELL,
        "wc_active": INDIC_WC_ACTIVE,
    }
    # Depth-bracket indicator slots (3 groups × 7 depths = 21 indicators, slots 9-29)
    _DEPTH_BASE = {"paren_d": 9, "sqbr_d": 14, "curly_d": 19}

    # Map depth tag names to indicator numbers
    _TAG_TO_INDIC_DEPTH: dict = {}
    for _pfx, _base in {"paren_d": 9, "sqbr_d": 14, "curly_d": 19}.items():
        for _i in range(5):
            _TAG_TO_INDIC_DEPTH[f"{_pfx}{_i}"] = _base + _i

    # Tags we silently ignore
    _IGNORED_TAGS = {"_pad", "bracket_match", "sel", "warn_tilde"}

    def tag_configure(self, tag, **kw):
        self._tag_cfg[tag] = kw
        # Push background color to the Scintilla indicator if this tag maps to one
        if self._hwnd:
            indic = self._TAG_TO_INDIC.get(tag)
            if indic is not None:
                bg = kw.get("background")
                if bg:
                    self._sci(SCI_INDICSETFORE, indic, self._color_to_bgr(bg))


    def tag_add(self, tag, index1, index2=None):
        if not self._hwnd or tag in self._IGNORED_TAGS:
            return
        indic = self._TAG_TO_INDIC.get(tag) or self._TAG_TO_INDIC_DEPTH.get(tag)
        if indic is None:
            return
        p1 = self._index_to_pos(index1)
        p2 = self._index_to_pos(index2) if index2 else p1 + 1
        if p2 > p1:
            # Find highlights override all other indicators — clear competing
            # background indicators from this range before filling, so they
            # don't bleed through even at high alpha.
            if indic in (INDIC_FIND_HL, INDIC_FIND_CUR):
                suppress = ([INDIC_WILDCARD, INDIC_ANGLE, INDIC_WC_ACTIVE]
                            + list(range(9, 24)))  # depth bracket slots 9-23
                for s_indic in suppress:
                    self._sci(SCI_SETINDICATORCURRENT, s_indic)
                    self._sci(SCI_INDICATORCLEARRANGE, p1, p2 - p1)
                # Track this range so _highlight_brackets can avoid repainting it
                self._find_hl_ranges.append((p1, p2))
            self._sci(SCI_SETINDICATORCURRENT, indic)
            self._sci(SCI_INDICSETVALUE, 1)
            self._sci(SCI_INDICATORFILLRANGE, p1, p2 - p1)
            # Track spell_err ranges in Python so tag_names can detect them
            # without relying on SCI_INDICATORVALUEAT (unreliable in NP++ Scintilla)
            if tag == "spell_err":
                self._tag_ranges.setdefault("spell_err", []).append((p1, p2))
            # Apply Scintilla text style for this tag if one is registered
            style = self._TAG_STYLE.get(tag)
            if style is not None:
                self._sci(SCI_STARTSTYLING, p1, 0xFF)
                self._sci(SCI_SETSTYLING, p2 - p1, style)
                self._styled_ranges.append((p1, p2, style))

    def tag_remove(self, tag, index1, index2=None):
        if not self._hwnd or tag in self._IGNORED_TAGS:
            return
        indic = self._TAG_TO_INDIC.get(tag) or self._TAG_TO_INDIC_DEPTH.get(tag)
        if indic is None:
            return
        p1 = self._index_to_pos(index1)
        p2 = self._index_to_pos(index2) if index2 else self._sci(SCI_GETLENGTH)
        if p2 > p1:
            self._sci(SCI_SETINDICATORCURRENT, indic)
            self._sci(SCI_INDICSETVALUE, 1)
            self._sci(SCI_INDICATORCLEARRANGE, p1, p2 - p1)
            # When find highlights are cleared, purge the protected ranges so
            # _highlight_brackets paints normally again.
            if indic in (INDIC_FIND_HL, INDIC_FIND_CUR):
                self._find_hl_ranges = [r for r in self._find_hl_ranges
                                        if not (r[0] >= p1 and r[1] <= p2)]
            if tag == "spell_err":
                # Clear all tracked spell ranges that fall within p1..p2
                existing = self._tag_ranges.get("spell_err", [])
                self._tag_ranges["spell_err"] = [r for r in existing
                                                  if not (r[0] >= p1 and r[1] <= p2)]
            # Revert text style to default (0) for this range
            style = self._TAG_STYLE.get(tag)
            if style is not None:
                self._sci(SCI_STARTSTYLING, p1, 0xFF)
                self._sci(SCI_SETSTYLING, p2 - p1, 0)
                self._styled_ranges = [(s,e,st) for s,e,st in self._styled_ranges
                                       if not (s >= p1 and e <= p2 and st == style)]

    def tag_raise(self, tag, *args):
        pass  # Scintilla indicator z-order is fixed by number

    def tag_ranges(self, tag):
        """Returns a flat tuple of index strings (start, end, start, end...)."""
        # We don't track ranges client-side for Scintilla indicators.
        # Return empty — callers that need this (find system) use their own list.
        return ()

    # ── Undo/redo ─────────────────────────────────────────────────────────────

    def edit_undo(self):
        if self._hwnd and self._sci(SCI_CANUNDO):
            self._sci(SCI_UNDO)

    def edit_redo(self):
        if self._hwnd and self._sci(SCI_CANREDO):
            self._sci(SCI_REDO)

    def edit_separator(self):
        if self._hwnd:
            self._sci(SCI_BEGINUNDOACTION)
            self._sci(SCI_ENDUNDOACTION)

    def edit_modified(self, value=None):
        if value is None:
            return bool(self._sci(SCI_GETMODIFY))
        if not value:
            self._sci(SCI_SETSAVEPOINT)
            self._modified_flag = False

    def edit_reset(self):
        if self._hwnd:
            self._sci(SCI_EMPTYUNDOBUFFER)

    # ── Geometry / info ───────────────────────────────────────────────────────

    def dlineinfo(self, index):
        """Return (x, y, width, height, baseline) for the line at index."""
        if not self._hwnd:
            return None
        try:
            pos  = self._index_to_pos(index)
            line = self._sci(SCI_LINEFROMPOSITION, pos)
            if line < 0:
                return None
            # SCI_POINTXFROMPOSITION=2164, SCI_POINTYFROMPOSITION=2165
            x = _user32.SendMessageW(self._hwnd, 2164, 0, pos)
            y = _user32.SendMessageW(self._hwnd, 2165, 0, pos)
            h = _user32.SendMessageW(self._hwnd, 2246, 0, line)  # SCI_TEXTHEIGHT
            if h <= 0:
                h = self._font_sz + 4  # fallback height
            return (x, y, self.winfo_width(), h, h - 2)
        except Exception:
            return None

    def bbox(self, index):
        if not self._hwnd:
            return None
        pos = self._index_to_pos(index)
        x = _user32.SendMessageW(self._hwnd, 2164, 0, pos)
        y = _user32.SendMessageW(self._hwnd, 2165, 0, pos)
        return (x, y, 8, 16)

    def winfo_height(self):
        return super().winfo_height()

    def winfo_ismapped(self):
        return super().winfo_ismapped()

    # ── Binding passthrough ───────────────────────────────────────────────────

    def bind(self, sequence, func=None, add=None):
        # Intercept virtual events — store in _virtual_cbs AND bind on frame
        # so event_generate("<<...>>") on self triggers the callback.
        if sequence and sequence.startswith("<<") and sequence.endswith(">>"):
            self._virtual_cbs.setdefault(sequence, [])
            if func:
                self._virtual_cbs[sequence].append(func)
            super().bind(sequence, func, add)
            return
        # Store double-click callbacks separately — Tk can't generate Double events
        if sequence == "<Double-Button-1>" and func:
            self._dbl_click_cbs.append(func)
            # Still bind to frame so real double-clicks on the frame work too
        # Ctrl+S / Ctrl+Shift+S and Ctrl+W / Ctrl+Shift+W:
        # Scintilla's HWND swallows these before Tk sees them, so frame bindings
        # never fire.  Route into _hotkey_cbs, called from _poll_sci_events.
        if sequence in ("<Control-s>", "<Control-S>") and func:
            shift = sequence == "<Control-S>"
            key = (True, shift, _VK_S)
            self._hotkey_cbs.setdefault(key, []).append(func)
            return  # do NOT fall through to super().bind — it's a no-op here
        if sequence in ("<Control-w>", "<Control-W>") and func:
            shift = sequence == "<Control-W>"
            key = (True, shift, _VK_W)
            self._hotkey_cbs.setdefault(key, []).append(func)
            return
        if sequence in ("<Control-f>", "<Control-F>") and func:
            key = (True, False, _VK_F)
            self._hotkey_cbs.setdefault(key, []).append(func)
            return
        if sequence in ("<Control-h>", "<Control-H>") and func:
            key = (True, False, _VK_H)
            self._hotkey_cbs.setdefault(key, []).append(func)
            return
        if sequence in ("<Control-n>", "<Control-N>") and func:
            key = (True, False, _VK_N)
            self._hotkey_cbs.setdefault(key, []).append(func)
            return
        super().bind(sequence, func, add)

    # ── Search ────────────────────────────────────────────────────────────────

    def search(self, pattern, index, stopindex=None, regexp=False,
               nocase=False, count=None, backwards=False):
        """Minimal search — returns index string or '' if not found."""
        if not self._hwnd:
            return ""
        flags = 0
        if not nocase:
            flags |= FIND_MATCHCASE
        if regexp:
            flags |= FIND_REGEXP
        start = self._index_to_pos(index)
        stop  = self._index_to_pos(stopindex) if stopindex else self._sci(SCI_GETLENGTH)
        needle = pattern.encode("utf-8") if not regexp else pattern.encode("utf-8")
        ttf = _TextToFind()
        ttf.chrg.cpMin  = start
        ttf.chrg.cpMax  = stop
        ttf.lpstrText   = needle
        pos = _user32.SendMessageW(self._hwnd, SCI_FINDTEXT, flags,
                                   _ptr(ctypes.pointer(ttf)))
        if pos < 0:
            return ""
        if count is not None:
            try:
                count[0] = ttf.chrgText.cpMax - ttf.chrgText.cpMin
            except Exception:
                pass
        return self._pos_to_index(pos)

    # ── Bracket highlighting (called by _apply_bracket_highlights) ────────────
    # The original app calls these; we implement them as Scintilla indicators.

    def update_brace_highlight(self, cursor_byte_pos):
        """Two jobs:
        1. Bold matched ( ) pair with INDIC_BRACE_BOLD
        2. Add SC_MARK_BACKGROUND line markers covering the unmatched span
           the cursor is inside (so empty lines get filled too).
        Depth/color stacking is already handled by _highlight_brackets.
        """
        if not self._hwnd:
            return
        total = self._sci(SCI_GETLENGTH)

        # Clear previous brace bold only — line markers handled by _highlight_brackets
        self._sci(SCI_SETINDICATORCURRENT, INDIC_BRACE_BOLD)
        self._sci(SCI_INDICATORCLEARRANGE, 0, total)

        if total == 0:
            return

        pos = cursor_byte_pos

        def byte_at(p):
            if 0 <= p < total:
                b = self._get_bytes(p, p + 1)
                return b if b else b""
            return b""

        anchor = -1
        if byte_at(pos) in (b"(", b")"):
            anchor = pos
        elif pos > 0 and byte_at(pos - 1) in (b"(", b")"):
            anchor = pos - 1

        if anchor < 0:
            return

        anchor_ch = byte_at(anchor)

        # Read full doc (cached)
        cache_key = (total, anchor)
        if getattr(self, "_brace_cache_key", None) == cache_key:
            raw = self._brace_cache_raw
        else:
            raw = self._get_bytes(0, total)
            self._brace_cache_key = cache_key
            self._brace_cache_raw = raw

        if anchor_ch == b"(":
            depth = 0
            match = -1
            for i in range(anchor, len(raw)):
                c = raw[i]
                if c == ord("("): depth += 1
                elif c == ord(")"):
                    depth -= 1
                    if depth == 0:
                        match = i
                        break
            if match >= 0:
                # Matched — bold both
                self._sci(SCI_SETINDICATORCURRENT, INDIC_BRACE_BOLD)
                self._sci(SCI_INDICSETVALUE, 1)
                self._sci(SCI_INDICATORFILLRANGE, anchor, 1)
                self._sci(SCI_INDICATORFILLRANGE, match, 1)
            else:
                pass  # markers already painted by _highlight_brackets
        else:
            depth = 0
            match = -1
            for i in range(anchor, -1, -1):
                c = raw[i]
                if c == ord(")"): depth += 1
                elif c == ord("("):
                    depth -= 1
                    if depth == 0:
                        match = i
                        break
            if match >= 0:
                # Matched — bold both
                self._sci(SCI_SETINDICATORCURRENT, INDIC_BRACE_BOLD)
                self._sci(SCI_INDICSETVALUE, 1)
                self._sci(SCI_INDICATORFILLRANGE, match, 1)
                self._sci(SCI_INDICATORFILLRANGE, anchor, 1)
            else:
                pass  # markers already painted by _highlight_brackets

    def _add_paren_markers(self, start_byte, end_byte, depth=0):
        """Add SC_MARK_BACKGROUND markers to lines in byte range at given depth."""
        marker = MARKER_PAREN_D[min(depth, 4)]
        total = self._sci(SCI_GETLENGTH)
        line_count = self._sci(2154)  # SCI_GETLINECOUNT
        start_line = self._sci(SCI_LINEFROMPOSITION, max(0, start_byte))
        end_line   = self._sci(SCI_LINEFROMPOSITION, min(max(total - 1, 0), end_byte))
        for ln in range(start_line, min(end_line + 1, line_count)):
            self._sci(SCI_MARKERADD, ln, marker)

    def _set_paren_span_markers(self, start_byte, end_byte):
        """Add MARKER_PAREN_SPAN to all lines in byte range."""
        self._add_paren_markers(start_byte, end_byte)


    def _highlight_brackets(self, content: str):
        """Apply all bracket/paren/angle/depth indicators using Scintilla."""
        if not self._hwnd:
            return
        content_bytes = content.encode("utf-8")
        content_byte_len = len(content_bytes)
        if content_byte_len <= 0:
            return
        sci_byte_len = self._sci(SCI_GETLENGTH)
        # Also get the full Scintilla buffer length so we can clear the full range

        def char_to_byte(char_pos):
            """Convert character offset to byte offset efficiently."""
            return len(content[:char_pos].encode("utf-8"))

        def fill_indic(indic, spans_char):
            """Fill indicator for a list of (start_char, end_char) spans,
            skipping any byte ranges protected by active find highlights."""
            self._sci(SCI_SETINDICATORCURRENT, indic)
            self._sci(SCI_INDICSETVALUE, 1)
            self._sci(SCI_INDICATORCLEARRANGE, 0, sci_byte_len)
            protected = self._find_hl_ranges  # [(byte_start, byte_end)]
            for cs, ce in spans_char:
                bs = char_to_byte(cs)
                be = char_to_byte(ce)
                if not (0 <= bs < be <= content_byte_len):
                    continue
                if not protected:
                    self._sci(SCI_INDICATORFILLRANGE, bs, be - bs)
                    continue
                # Fill only the sub-ranges not covered by a find highlight
                cursor = bs
                for pb, pe in sorted(protected):
                    if pe <= cursor or pb >= be:
                        continue  # no overlap
                    if cursor < pb:
                        self._sci(SCI_INDICATORFILLRANGE, cursor, pb - cursor)
                    cursor = max(cursor, pe)
                if cursor < be:
                    self._sci(SCI_INDICATORFILLRANGE, cursor, be - cursor)

        # ── Angle brackets ────────────────────────────────────────────────────
        fill_indic(INDIC_ANGLE,
                   [(m.start(), m.end()) for m in re.finditer(r"<[^<>\n]*>", content)])

        # ── Depth-colored brackets: paren, sqbr, curly ───────────────────────
        for _m in MARKER_PAREN_D:
            self._sci(SCI_MARKERDELETEALL, _m)  # clear all depth markers before repainting
        for open_ch, close_ch, base_indic in (
            ("(", ")", 9),    # paren_d0-4
            ("[", "]", 14),   # sqbr_d0-4
            ("{", "}", 19),   # curly_d0-4
        ):
            spans_by_depth = [[] for _ in range(5)]
            stack = []
            unmatched_close_spans = []  # (0, close_pos+1) for unmatched closes
            open_stack = []  # track all opens to find unmatched ones
            for i, ch in enumerate(content):
                if ch == open_ch:
                    stack.append(i)
                elif ch == close_ch:
                    if stack:
                        start = stack.pop()
                        depth = min(len(stack), 4)
                        spans_by_depth[depth].append((start, i + 1))
                    else:
                        # Unmatched close — span from 0 to here
                        unmatched_close_spans.append((0, i + 1))
            # Unmatched opens — highlight to end of real content
            # stack contains positions in order pushed; deeper nesting = higher index
            unmatched_open_spans = []  # (start_char, depth)
            for idx, start in enumerate(stack):
                depth = min(idx, 4)
                spans_by_depth[depth].append((start, len(content)))
                unmatched_open_spans.append((start, depth))
            for depth, spans in enumerate(spans_by_depth):
                if spans:
                    fill_indic(base_indic + depth, spans)
                else:
                    self._sci(SCI_SETINDICATORCURRENT, base_indic + depth)
                    self._sci(SCI_INDICATORCLEARRANGE, 0, sci_byte_len)
            # Add line markers for unmatched paren spans (fills empty lines)
            if open_ch == "(":
                for start, depth in unmatched_open_spans:
                    bs = char_to_byte(start)
                    self._add_paren_markers(bs, content_byte_len, depth)
                for cs, ce in unmatched_close_spans:
                    be = char_to_byte(ce)
                    self._add_paren_markers(0, be, 0)
