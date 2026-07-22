import os
import json
import threading
import webbrowser
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import font as tkfont

from ..constants import (APP_AI_VOICES_DIR, APP_PIPER_DIR, APP_PIPER_BIN_DIR,
                          APP_PIPER_PRESETS_DIR, APP_PIPER_VOICES_DIR,
                          APP_KOKORO_DIR, APP_KOKORO_CUSTOM_DIR, APP_SETTINGS_PATH,
                          APP_SAPI_PRESETS_DIR)
from ..kokoro_catalog import (KOKORO_VOICES, KOKORO_GITHUB_API,
                               KOKORO_MODEL_FILE, KOKORO_VOICES_FILE,
                               load_voices_from_file)
from ..window_geometry import apply_window_geometry
from ..screen_capture import get_dpi_scale

PIPER_VOICES_JSON_URL = "https://huggingface.co/rhasspy/piper-voices/resolve/main/voices.json"
PIPER_HF_BASE        = "https://huggingface.co/rhasspy/piper-voices/resolve/main"


class _ManageWindow:
    """Manage/download popup for Piper engine or Kokoro model."""

    def __init__(self, parent, title, info_lines, size_text,
                 status_text, action_label, on_download, on_cancel,
                 icon_path=None, repo_url=None, repo_urls=None,
                 note_lines=None, delete_label=None, on_delete=None,
                 status_color="#333333"):
        self.cancelled = False
        self._downloading = False
        self._on_download = on_download
        self._on_cancel = on_cancel
        self._on_delete = on_delete

        self.win = tk.Toplevel(parent)
        self.win.title(title)
        self.win.resizable(False, False)
        self.win.grab_set()
        self.win.protocol("WM_DELETE_WINDOW", self._close)
        if icon_path and os.path.exists(icon_path):
            try:
                self.win.iconbitmap(icon_path)
            except Exception:
                pass

        frame = tk.Frame(self.win, padx=24, pady=16)
        frame.pack(fill=tk.BOTH, expand=True)

        for line in info_lines:
            tk.Label(frame, text=line, font=("Helvetica", 10, "bold"),
                     anchor='w').pack(anchor='w')
        if note_lines:
            for line in note_lines:
                tk.Label(frame, text=line, font=("Helvetica", 9),
                         fg="#555555", anchor='w').pack(anchor='w')
        tk.Label(frame, text=f"Download size:  {size_text}",
                 font=("Helvetica", 9), fg="#555555").pack(anchor='w', pady=(4, 2))
        self._static_status_label = tk.Label(frame, text=f"Status:  {status_text}",
                 font=("Helvetica", 9), fg=status_color)
        self._static_status_label.pack(anchor='w', pady=(0, 14))

        self.progress = ttk.Progressbar(frame, orient=tk.HORIZONTAL,
                                        mode='determinate', length=420)
        self.progress.pack(fill=tk.X, pady=(0, 4))

        links = repo_urls if repo_urls else ([(("HuggingFace:  " if "huggingface" in repo_url else "GitHub:  "), repo_url)] if repo_url else [])
        for prefix, url in links:
            repo_row = tk.Frame(frame)
            repo_row.pack(anchor='w', pady=(2, 0))
            tk.Label(repo_row, text=prefix, font=("Helvetica", 9),
                     fg="#555555").pack(side=tk.LEFT)
            link = tk.Label(repo_row, text=url, font=("Helvetica", 9),
                            fg="#0066cc", cursor="hand2")
            link.pack(side=tk.LEFT)
            link.bind("<Button-1>", lambda _e, u=url: webbrowser.open(u))

        self._dl_label = tk.Label(frame, text="", font=("Helvetica", 9),
                                  fg="#444444", anchor='w')
        self._dl_label.pack(anchor='w', pady=(4, 10))

        btn_row = tk.Frame(frame)
        btn_row.pack()
        self._action_btn = tk.Button(btn_row, text=action_label, width=22,
                                     command=self._on_action_click)
        self._action_btn.pack(side=tk.LEFT, padx=(0, 10))
        self._cancel_btn = tk.Button(btn_row, text="Cancel", width=10,
                                     command=self._close)
        self._cancel_btn.pack(side=tk.LEFT)

        if on_delete:
            self._delete_btn = tk.Button(btn_row, text=delete_label or "Delete",
                                         width=14, command=self._on_delete_click)
            self._delete_btn.pack(side=tk.LEFT, padx=(10, 0))

    def _on_action_click(self):
        self._downloading = True
        self._action_btn.config(state=tk.DISABLED)
        self._dl_label.config(text="Starting...", fg="#444444")
        if self._on_download:
            self._on_download()

    def _on_delete_click(self):
        if self._on_delete:
            self._on_delete()

    def set_progress(self, pct):
        if self.win.winfo_exists():
            self.progress["value"] = pct

    def set_status(self, text, color="black"):
        if self.win.winfo_exists():
            self._dl_label.config(text=text, fg=color)

    def set_static_status(self, text, color="#333333"):
        if self.win.winfo_exists():
            self._static_status_label.config(text=f"Status:  {text}", fg=color)

    def mark_done(self):
        self._downloading = False
        if self.win.winfo_exists():
            self._cancel_btn.config(text="Close", command=self.close)

    def _close(self):
        self.cancelled = True
        if self._downloading and self._on_cancel:
            self._on_cancel()
        if self.win.winfo_exists():
            self.win.grab_release()
            self.win.destroy()

    def close(self):
        if self.win.winfo_exists():
            self.win.grab_release()
            self.win.destroy()


class _VoiceListbox:
    """Treeview-backed widget that provides a Listbox-compatible API with per-item bold support."""

    def __init__(self, parent, font=("Helvetica", 9), **kwargs):
        self._font_normal = font
        self._font_bold = (font[0], font[1], "bold")
        style_name = "VoiceList.Treeview"
        s = ttk.Style()
        # Derive rowheight from the actual font metrics rather than a fixed pixel
        # count - a hardcoded value doesn't scale with Windows display scaling
        # (e.g. 300%), causing the text to be clipped inside too-short rows.
        row_height = tkfont.Font(font=font).metrics("linespace") + 6
        s.configure(style_name, font=font, rowheight=row_height)
        # Remove the outer border so it looks like a plain listbox.
        # The layout element name varies by theme, so guard the call.
        try:
            s.layout(style_name, [("Treeview.treearea", {"sticky": "nswe"})])
        except Exception:
            pass
        self._tree = ttk.Treeview(parent, show="tree", selectmode="extended",
                                   style=style_name)
        try:
            self._tree.column("#0", stretch=True, minwidth=0)
        except Exception:
            pass
        self._tree.tag_configure("normal",  font=font)
        self._tree.tag_configure("bold",    font=self._font_bold)
        self._tree.tag_configure("problem", font=font, foreground="#999999")
        self._tree.tag_configure("problem_bold", font=self._font_bold, foreground="#999999")
        self.yview = self._tree.yview

    def configure(self, **kwargs):
        if "yscrollcommand" in kwargs:
            self._tree.configure(yscrollcommand=kwargs.pop("yscrollcommand"))
        if kwargs:
            self._tree.configure(**kwargs)

    def pack(self, **kwargs):
        self._tree.pack(**kwargs)

    def insert(self, _index, text, bold=False, problem=False):
        if problem:
            tag = "problem_bold" if bold else "problem"
        else:
            tag = "bold" if bold else "normal"
        self._tree.insert("", "end", text=text, tags=(tag,))

    def identify_row(self, y):
        return self._tree.identify_row(y)

    def index_of_iid(self, iid):
        children = self._tree.get_children()
        return children.index(iid) if iid in children else None

    def delete(self, first, last=None):
        children = self._tree.get_children()
        if not children:
            return
        if last is None:
            if first < len(children):
                self._tree.delete(children[first])
        elif last in (tk.END, "end"):
            self._tree.delete(*children[first:])
        else:
            self._tree.delete(*children[first:last + 1])

    def curselection(self):
        sel_iids = self._tree.selection()
        children  = self._tree.get_children()
        return tuple(i for i, iid in enumerate(children) if iid in sel_iids)

    def clear_selection(self):
        sel = self._tree.selection()
        if sel:
            self._tree.selection_remove(*sel)

    def bind(self, event, handler, add=None):
        mapped = "<<TreeviewSelect>>" if event == "<<ListboxSelect>>" else event
        kw = {"add": add} if add is not None else {}
        self._tree.bind(mapped, handler, **kw)


class AIVoiceDownloadWindow:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self._voice_catalog = {}   # key -> info dict (fetched from HuggingFace)

        self.window = tk.Toplevel(parent)
        self.window.withdraw()
        self.window.title("Voice Manager")
        self.window.resizable(True, True)
        apply_window_geometry(self.window, 'voice_manager', 1100, 710)
        self.window.minsize(*[int(v * get_dpi_scale()) for v in (400, 300)])

        try:
            icon_path = os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico')
            if os.path.exists(icon_path):
                self.window.iconbitmap(icon_path)
        except Exception:
            pass

        os.makedirs(APP_AI_VOICES_DIR, exist_ok=True)
        os.makedirs(APP_PIPER_VOICES_DIR, exist_ok=True)
        os.makedirs(APP_PIPER_BIN_DIR, exist_ok=True)
        os.makedirs(APP_PIPER_PRESETS_DIR, exist_ok=True)
        self._setup_ui()
        self._refresh_engine_status()
        self._apply_cached_update_info()
        self._fetch_voice_catalog()
        self.window.deiconify()

        # After grab_set/grab_release cycles (e.g. auto-read area selection) tkinter's
        # internal focus state can get corrupted so Entry/Text widgets stop accepting
        # keyboard input even when the window is visible and focused.
        # Fix 1: force OS-level focus whenever this Toplevel gets activated.
        self.window.bind('<FocusIn>', self._on_window_focus_in)
        # Fix 2: force widget-level focus on any click into an Entry or Text widget,
        # bypassing the normal focus_set() that can silently fail after a grab cycle.
        self.window.bind('<Button-1>', self._on_window_click, add='+')

    def _on_window_focus_in(self, event):
        if event.widget is self.window:
            self.window.focus_force()

    def _on_window_click(self, event):
        if isinstance(event.widget, (tk.Entry, tk.Text)):
            event.widget.focus_force()

    # ---------------------------------------------------------------- UI setup

    def _make_scrollable(self, parent, horizontal=False):
        """Wrap parent in a canvas+scrollbar(s) and return an inner frame to build content into.

        By default only vertical scrolling is enabled and the inner frame's
        width is kept locked to the canvas viewport width, so `fill=X`
        children behave as if they were packed directly into `parent`. Pass
        horizontal=True for tabs that pack several fixed-width columns
        side by side and can also overflow sideways - the inner frame then
        keeps its own natural (unclamped) width and a horizontal scrollbar
        is added too.
        """
        canvas = tk.Canvas(parent, highlightthickness=0)
        vsb = tk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        if horizontal:
            hsb = tk.Scrollbar(parent, orient=tk.HORIZONTAL, command=canvas.xview)
            canvas.configure(xscrollcommand=hsb.set)
            hsb.pack(side=tk.BOTTOM, fill=tk.X)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = tk.Frame(canvas)
        win_id = canvas.create_window((0, 0), window=inner, anchor='nw')

        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        if not horizontal:
            canvas.bind('<Configure>', lambda e: canvas.itemconfig(win_id, width=e.width))

        def _scroll(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind('<Enter>', lambda e: canvas.bind_all('<MouseWheel>', _scroll))
        canvas.bind('<Leave>', lambda e: canvas.unbind_all('<MouseWheel>'))

        return inner

    def _setup_ui(self):
        main = tk.Frame(self.window, padx=10, pady=10)
        main.pack(fill=tk.BOTH, expand=True)

        tk.Label(main, text="Voice Manager",
                 font=("Helvetica", 12, "bold")).pack(pady=(0, 2))

        # Shared maps
        self._available_map = {}
        self._installed_map = {}

        # Notebook tabs
        notebook = ttk.Notebook(main)
        notebook.pack(fill=tk.BOTH, expand=True)

        # ── Voice Selection tab (first) ───────────────────────────────
        voices_tab = tk.Frame(notebook, padx=8, pady=8)
        notebook.add(voices_tab, text="  Voice Dropdown  ")
        self._build_voices_tab(voices_tab)

        # ── Piper tab ────────────────────────────────────────────────
        piper_tab = tk.Frame(notebook, padx=8, pady=8)
        notebook.add(piper_tab, text="  Piper Voices  ")
        self._build_piper_tab(piper_tab)

        # ── Kokoro tab ───────────────────────────────────────────────
        kokoro_tab = tk.Frame(notebook, padx=8, pady=2)
        notebook.add(kokoro_tab, text="  Kokoro Voices  ")
        self._build_kokoro_tab(kokoro_tab)

        # ── Windows (SAPI) tab ───────────────────────────────────────
        sapi_tab = tk.Frame(notebook, padx=8, pady=8)
        notebook.add(sapi_tab, text="  Windows (SAPI)  ")
        self._build_sapi_tab(sapi_tab)

        # Bottom buttons
        ctrl = tk.Frame(main)
        ctrl.pack(fill=tk.X)
        tk.Button(ctrl, text="Open Voices Folder",
                  command=self._open_folder, width=18).pack(side=tk.LEFT)
        tk.Button(ctrl, text="Close",
                  command=self.window.destroy, width=10).pack(side=tk.RIGHT)

    def _build_voices_tab(self, parent):
        self._voices_playing = False
        self._voices_play_btn: tk.Button
        self._voices_play_token = 0   # incremented on every new play; stale callbacks ignored

        # ── Download / Install section ────────────────────────────────
        install_frame = tk.Frame(parent)
        install_frame.pack(fill=tk.X, pady=(0, 2))

        piper_row = tk.Frame(install_frame)
        piper_row.pack(fill=tk.X, pady=(0, 2))
        piper_left = tk.Frame(piper_row)
        piper_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(piper_left, text="AI Voice Piper:", width=15, anchor='w',
                 font=("Helvetica", 9, "bold")).pack(side=tk.LEFT)
        self.engine_status_label = tk.Label(piper_left, text="Built-in",
                                             fg="green", font=("Helvetica", 9))
        self.engine_status_label.pack(side=tk.LEFT, padx=(4, 0))
        piper_right = tk.Frame(piper_row)
        piper_right.pack(side=tk.RIGHT)
        self.engine_btn = tk.Button(piper_right, text="GitHub ↗",
                                     command=self._open_engine_manage, width=10)
        self.engine_btn.pack(side=tk.LEFT)

        kokoro_row = tk.Frame(install_frame)
        kokoro_row.pack(fill=tk.X)
        kokoro_left = tk.Frame(kokoro_row)
        kokoro_left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Label(kokoro_left, text="AI Voice Kokoro:", width=15, anchor='w',
                 font=("Helvetica", 9, "bold")).pack(side=tk.LEFT)
        self.kokoro_status_label = tk.Label(kokoro_left, text="Not installed",
                                             fg="#999999", font=("Helvetica", 9))
        self.kokoro_status_label.pack(side=tk.LEFT, padx=(4, 0))
        kokoro_right = tk.Frame(kokoro_row)
        kokoro_right.pack(side=tk.RIGHT)
        self.kokoro_download_btn = tk.Button(kokoro_right, text="Manage",
                                              command=self._open_kokoro_manage, width=10)
        self.kokoro_download_btn.pack(side=tk.LEFT)

        ttk.Separator(parent, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(4, 4))

        # ── main layout: place-based so proportions are truly enforced ──
        body = tk.Frame(parent)
        body.pack(fill=tk.BOTH, expand=True)

        # Left (0–46%) — available voices
        left_frame = tk.LabelFrame(body, text="Available Voices", padx=5, pady=5)
        left_frame.place(relx=0.0, rely=0.0, relwidth=0.46, relheight=1.0)

        self._all_voices_list = _VoiceListbox(left_frame, font=("Helvetica", 9))
        sb_l = tk.Scrollbar(left_frame, orient=tk.VERTICAL,
                             command=self._all_voices_list.yview)
        self._all_voices_list.configure(yscrollcommand=sb_l.set)
        sb_l.pack(side=tk.RIGHT, fill=tk.Y)
        self._all_voices_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Middle (46–56%) — Add/Remove buttons, vertically centred
        mid = tk.Frame(body)
        mid.place(relx=0.46, rely=0.0, relwidth=0.10, relheight=1.0)
        tk.Frame(mid).pack(side=tk.TOP, expand=True, fill=tk.BOTH)
        tk.Button(mid, text="Add →", width=10,
                  command=self._voices_add).pack(side=tk.TOP, pady=4)
        tk.Button(mid, text="← Remove", width=10,
                  command=self._voices_remove).pack(side=tk.TOP, pady=4)
        tk.Frame(mid).pack(side=tk.TOP, expand=True, fill=tk.BOTH)

        # Right (56–100%) — test box, play button, selection list
        right_frame = tk.Frame(body)
        right_frame.place(relx=0.56, rely=0.0, relwidth=0.44, relheight=1.0)

        self._voices_status_label = tk.Label(right_frame, text="",
                                              font=("Helvetica", 9, "bold"),
                                              fg="black", anchor="w")
        self._voices_status_label.pack(fill=tk.X, pady=(0, 0))

        self._voices_selected_name_label = tk.Label(right_frame, text="Selected Voice: —",
                                                     font=("Helvetica", 9, "bold"), fg="black",
                                                     anchor="w", width=1)
        self._voices_selected_name_label.pack(fill=tk.X, pady=(2, 4))

        tk.Label(right_frame, text="Test text:", anchor="w",
                 font=("Helvetica", 9)).pack(fill=tk.X)

        self._voices_test_widget = tk.Text(right_frame, font=("Helvetica", 9),
                                            height=4, wrap="word", width=1)
        self._voices_test_widget.insert("1.0",
            "The quick brown fox jumps over the lazy dog. Testing one two three.")
        self._voices_test_widget.pack(fill=tk.X, pady=(0, 2))
        self._voices_test_widget.bind('<Button-1>',
            lambda _: (self.window.focus_force(), self._voices_test_widget.focus_set()), add='+')

        tk.Label(right_frame,
                 text="Select a voice then press Play.",
                 font=("Helvetica", 8), fg="#777777", anchor="w").pack(fill=tk.X, pady=(0, 4))

        btn_row = tk.Frame(right_frame)
        btn_row.pack(anchor="w", pady=(0, 6))
        self._voices_play_btn = tk.Button(btn_row, text="▶ Play",
                                           width=12, command=self._voices_play)
        self._voices_play_btn.pack(side=tk.LEFT)
        total = len(self._TEST_TEXTS)
        self._voices_next_btn = tk.Button(btn_row,
                                          text=f"1/{total}  Next Text",
                                          width=14,
                                          command=self._voices_next_test_text)
        self._voices_next_btn.pack(side=tk.LEFT, padx=(8, 0))

        selected_voices_label = tk.Label(right_frame, text="Selected Voices (shown in main window voice dropdown):",
                 anchor="w", justify="left", font=("Helvetica", 12, "bold"))
        selected_voices_label.pack(fill=tk.X)
        # right_frame is placed at a relative width (proportional layout), so
        # a fixed wraplength would be wrong at other window sizes - keep it
        # tied to the label's actual current width instead.
        selected_voices_label.bind('<Configure>', lambda e: selected_voices_label.config(wraplength=e.width))
        self._selected_voices_list = _VoiceListbox(right_frame, font=("Helvetica", 9))
        sb_r = tk.Scrollbar(right_frame, orient=tk.VERTICAL,
                             command=self._selected_voices_list.yview)
        self._selected_voices_list.configure(yscrollcommand=sb_r.set)
        sb_r.pack(side=tk.RIGHT, fill=tk.Y)
        self._selected_voices_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._all_voices_list.bind('<<ListboxSelect>>', self._voices_on_list_select)
        self._selected_voices_list.bind('<<ListboxSelect>>', self._voices_on_list_select)
        self._selected_voices_list.bind('<Button-1>', self._voices_on_selected_click, add='+')

        # Populate both lists
        self._voices_tab_data = []   # list of (label, engine, key) for all voices
        self._selected_voices_data = []  # same structure for selected list
        self._populate_voices_lists()
        self._start_sapi_worker()
        self.window.protocol("WM_DELETE_WINDOW",
                             lambda: (self._sapi_q.put(None), self.window.destroy()))

    # Engine types that are user-created custom presets/voices
    _CUSTOM_ENGINES = {"sapi_preset", "piper_preset", "kokoro_custom"}

    def _enumerate_all_voices(self):
        """Return list of (label, engine, key) for every available voice.
        Custom voices are always placed at the end of the list."""
        regular = []
        custom  = []

        # SAPI — try COM first, fall back to winreg if NVDA or another app blocks GetVoices()
        sapi_names = []
        try:
            import win32com.client
            sapi = win32com.client.Dispatch("SAPI.SpVoice")
            for token in sapi.GetVoices():
                try:
                    sapi_names.append(token.GetDescription())
                except Exception:
                    pass
        except Exception:
            pass
        if not sapi_names:
            import winreg as _wr
            for _sub in (r"SOFTWARE\Microsoft\Speech\Voices\Tokens",
                         r"SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens"):
                try:
                    _root = _wr.OpenKey(_wr.HKEY_LOCAL_MACHINE, _sub)
                    _i = 0
                    while True:
                        try:
                            _tok = _wr.EnumKey(_root, _i); _i += 1
                        except OSError:
                            break
                        _desc = None
                        try:
                            _tk = _wr.OpenKey(_root, _tok)
                            for _s in ("409", ""):
                                try:
                                    if _s:
                                        _sk = _wr.OpenKey(_tk, _s)
                                        _desc, _ = _wr.QueryValueEx(_sk, ""); _wr.CloseKey(_sk)
                                    else:
                                        _desc, _ = _wr.QueryValueEx(_tk, "")
                                    if _desc: break
                                except Exception: pass
                            if not _desc:
                                try:
                                    _ak = _wr.OpenKey(_tk, "Attributes")
                                    _desc, _ = _wr.QueryValueEx(_ak, "Name"); _wr.CloseKey(_ak)
                                except Exception: pass
                            _wr.CloseKey(_tk)
                        except Exception: pass
                        if _desc and _desc not in sapi_names:
                            sapi_names.append(_desc)
                    _wr.CloseKey(_root)
                except Exception:
                    pass
        for name in sapi_names:
            regular.append((f"[SAPI] {name}", "sapi", name))

        # Piper — installed .onnx files (only shown when the engine binary is present)
        try:
            from ..constants import APP_PIPER_VOICES_DIR as _piper_v_dir, APP_PIPER_PRESETS_DIR
            if self._piper_exe_exists():
                if os.path.isdir(_piper_v_dir):
                    for fname in sorted(os.listdir(_piper_v_dir)):
                        if fname.endswith('.onnx') and not fname.startswith('_'):
                            key = fname[:-5]
                            label = self._piper_key_label(key)
                            regular.append((f"[Piper] {label}", "piper", key))
                if os.path.isdir(APP_PIPER_PRESETS_DIR):
                    for fname in sorted(os.listdir(APP_PIPER_PRESETS_DIR)):
                        if fname.endswith('.json'):
                            name = fname[:-5]
                            custom.append((f"[Piper] {name} (custom)", "piper_preset", name))
        except Exception:
            pass

        # Kokoro — only shown when the model file is actually installed
        try:
            from ..constants import APP_KOKORO_DIR
            from ..kokoro_catalog import load_voices_from_file
            voices_bin = os.path.join(APP_KOKORO_DIR, 'voices-v1.0.bin')
            if os.path.exists(voices_bin):
                vdict = load_voices_from_file(voices_bin)
                for voice_id, (display, _lang) in vdict.items():
                    regular.append((f"[Kokoro] {display}", "kokoro", voice_id))
                from ..constants import APP_KOKORO_CUSTOM_DIR
                if os.path.isdir(APP_KOKORO_CUSTOM_DIR):
                    for fname in sorted(os.listdir(APP_KOKORO_CUSTOM_DIR)):
                        if fname.endswith('.npy'):
                            name = fname[:-4]
                            custom.append((f"[Kokoro] {name}", "kokoro_custom", name))
        except Exception:
            pass

        # SAPI custom presets — appended after Kokoro so all customs are grouped at bottom
        try:
            if os.path.isdir(APP_SAPI_PRESETS_DIR):
                for fname in sorted(os.listdir(APP_SAPI_PRESETS_DIR)):
                    if fname.endswith('.json'):
                        name = fname[:-5]
                        custom.append((f"{name} (custom)", "sapi_preset", name))
        except Exception:
            pass

        return regular + custom

    def _rebuild_selected_list(self):
        """Redraw the selected voices listbox using each voice's original available-list number."""
        self._selected_voices_list.delete(0, tk.END)
        nums = getattr(self, '_voice_numbers', {})
        fallback_nums = getattr(self, '_fallback_voice_numbers', {})
        for label, eng, key in self._selected_voices_data:
            problem = (eng, key) in fallback_nums
            num  = nums.get((eng, key), fallback_nums.get((eng, key), "?"))
            bold = eng in self._CUSTOM_ENGINES
            suffix = "  (?)" if problem else ""
            self._selected_voices_list.insert(tk.END, f"{num}.  {label}{suffix}", bold=bold, problem=problem)

    # ── settings persistence ──────────────────────────────────────────

    def _voices_save(self):
        """Save the selected voices list to gametextreader_settings.json."""
        try:
            settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            nums = getattr(self, '_voice_numbers', {})
            fallback_nums = getattr(self, '_fallback_voice_numbers', {})
            settings['voice_filter'] = [
                {"engine": e, "key": k, "num": nums.get((e, k), fallback_nums.get((e, k), i + 1))}
                for i, (_label, e, k) in enumerate(self._selected_voices_data)
            ]
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
        except Exception as ex:
            print(f"[ERROR] Could not save voice filter: {ex}")

    def _voices_load_saved(self):
        """Return saved voice_filter as a set of (engine, key) tuples, or None if not saved."""
        try:
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                raw = settings.get('voice_filter')
                if isinstance(raw, list) and len(raw) > 0:
                    return {(item['engine'], item['key']) for item in raw}
        except Exception:
            pass
        return None

    def _voices_load_saved_ordered(self):
        """Return saved voice_filter as an ordered list of (engine, key, num) tuples, or None if not saved.
        `num` is the number that was saved alongside the entry (or None)."""
        try:
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                raw = settings.get('voice_filter')
                if isinstance(raw, list) and len(raw) > 0:
                    return [(item['engine'], item['key'], item.get('num')) for item in raw]
        except Exception:
            pass
        return None

    def _populate_voices_lists(self):
        all_voices = self._enumerate_all_voices()
        self._voices_tab_data = all_voices

        # Assign a permanent 1-based number to every voice (order = enumeration order).
        # This number travels with the voice regardless of which list it lives in.
        self._voice_numbers = {(e, k): i + 1 for i, (_, e, k) in enumerate(all_voices)}

        saved         = self._voices_load_saved()           # set of (engine, key) or None
        saved_ordered = self._voices_load_saved_ordered()   # ordered [(engine, key, num)] or None

        # Numbers to fall back on for saved voices that can't be matched to anything
        # currently enumerable (e.g. cloud "Online (Natural)" voices that have no
        # registry token and so never appear in `all_voices`). Without this they'd
        # show as "?." even though we know what number they were saved under.
        self._fallback_voice_numbers = {}

        if saved is not None and saved_ordered is not None:
            all_voices_map = {(e, k): l for l, e, k in all_voices}

            def _sapi_base(name):
                return name.split(" - ")[0].strip().lower()
            sapi_by_base = {}
            for (eng, key), lbl in all_voices_map.items():
                if eng == "sapi":
                    sapi_by_base.setdefault(_sapi_base(key), (key, lbl))

            # Preserve saved order; keep voices even if not currently enumerable (e.g. SAPI broken)
            selected = []
            used_keys = set()
            for engine, key, saved_num in saved_ordered:
                if (engine, key) in all_voices_map:
                    label = all_voices_map[(engine, key)]
                    selected.append((label, engine, key))
                    used_keys.add((engine, key))
                elif engine == "sapi":
                    # Fuzzy match: find a current SAPI voice with the same base name
                    base = _sapi_base(key)
                    match = sapi_by_base.get(base)
                    if not match:
                        for cur_base, cur_val in sapi_by_base.items():
                            if base in cur_base or cur_base in base:
                                match = cur_val
                                break
                    if match:
                        matched_key, matched_label = match
                        selected.append((matched_label, engine, matched_key))
                        used_keys.add((engine, matched_key))
                    else:
                        selected.append((f"[SAPI] {key}", engine, key))
                        used_keys.add((engine, key))
                        if saved_num is not None:
                            self._fallback_voice_numbers[(engine, key)] = saved_num
                else:
                    selected.append((f"[{engine.upper()}] {key}", engine, key))
                    used_keys.add((engine, key))
                    if saved_num is not None:
                        self._fallback_voice_numbers[(engine, key)] = saved_num
            available = [(l, e, k) for l, e, k in all_voices if (e, k) not in used_keys]

            # Always keep the Selected list in numbered order, regardless of the order
            # it happened to be saved in.
            _nums, _fallback = self._voice_numbers, self._fallback_voice_numbers
            _orig_order = [(e, k) for _l, e, k in selected]
            selected.sort(key=lambda x: _nums.get((x[1], x[2]), _fallback.get((x[1], x[2]), float('inf'))))
            _reordered = [(e, k) for _l, e, k in selected] != _orig_order
        else:
            selected   = [(l, e, k) for l, e, k in all_voices if e == "sapi"]
            available  = [(l, e, k) for l, e, k in all_voices if e != "sapi"]
            _reordered = False

        self._voices_available_data = available
        self._all_voices_list.delete(0, tk.END)
        for label, eng, key in available:
            num  = self._voice_numbers.get((eng, key), "?")
            bold = eng in self._CUSTOM_ENGINES
            self._all_voices_list.insert(tk.END, f"{num}.  {label}", bold=bold)

        self._selected_voices_data = selected
        self._rebuild_selected_list()
        if _reordered:
            self._voices_save()

    def _voices_add(self):
        sel = self._all_voices_list.curselection()
        if not sel:
            return
        # Remove highest indices first so lower indices stay valid
        for idx in sorted(sel, reverse=True):
            entry = self._voices_available_data.pop(idx)
            self._all_voices_list.delete(idx)
            self._selected_voices_data.append(entry)
        # Keep the Selected list in numbered order rather than insertion order
        nums = getattr(self, '_voice_numbers', {})
        fallback_nums = getattr(self, '_fallback_voice_numbers', {})
        self._selected_voices_data.sort(
            key=lambda x: nums.get((x[1], x[2]), fallback_nums.get((x[1], x[2]), float('inf')))
        )
        self._rebuild_selected_list()
        self._voices_save()
        try:
            self.app.refresh_voice_dropdowns()
        except Exception:
            pass

    def _voices_remove(self):
        sel = self._selected_voices_list.curselection()
        if not sel:
            return
        for idx in sorted(sel, reverse=True):
            entry = self._selected_voices_data.pop(idx)
            self._voices_available_data.append(entry)
        self._rebuild_selected_list()
        # Re-sort available list by original voice number so order stays consistent
        nums = getattr(self, '_voice_numbers', {})
        self._voices_available_data.sort(key=lambda x: nums.get((x[1], x[2]), float('inf')))
        self._all_voices_list.delete(0, tk.END)
        for lbl, eng, key in self._voices_available_data:
            num  = nums.get((eng, key), "?")
            bold = eng in self._CUSTOM_ENGINES
            self._all_voices_list.insert(tk.END, f"{num}.  {lbl}", bold=bold)
        self._voices_save()
        try:
            self.app.refresh_voice_dropdowns()
        except Exception:
            pass

    _ENGINE_COLORS = {
        "sapi":         "#226600",
        "sapi_preset":  "#226600",
        "piper":        "#0055aa",
        "piper_preset": "#0055aa",
        "kokoro":       "#7700aa",
        "kokoro_custom":"#7700aa",
    }

    _TEST_TEXTS = [
        # ── Short sentences ───────────────────────────────────────────
        "The quick brown fox jumps over the lazy dog.",
        "Testing one, two, three. How does this voice sound to you?",
        "Good morning. Please proceed to gate number seven immediately.",
        "She smiled and nodded once before quietly leaving the room.",
        "System error detected on startup. Please contact your administrator.",
        # ── Longer passages ───────────────────────────────────────────
        "Every morning the old train passed quietly through the snow-covered valley. "
        "The passengers sat in silence, watching the frozen landscape drift past the frosted windows. "
        "Please verify your email address carefully before continuing to the next step.",
        "Dynamic range compression changes how quiet and loud sounds behave in a recording. "
        "Her pronunciation became noticeably clearer after weeks of practicing difficult vowel sounds daily. "
        "The software update will begin downloading shortly. Please do not close this window.",
        "Testing emotional range: happy and carefree, deeply serious, wildly excited, anxious and nervous, "
        "and finally perfectly calm. Can your voice model handle shifting tone without sounding robotic? "
        "This is an important test for expressive speech synthesis.",
        "The ancient map showed three possible routes through the forest, but only one was safe after dark. "
        "Torches lined the path ahead, flickering in the cold wind as the group moved cautiously forward. "
        "No one spoke. The only sound was the crunch of frozen leaves underfoot.",
        "Welcome back, Commander. Your last mission was a success, but the war is far from over. "
        "New intelligence reports suggest the enemy has repositioned their forces along the northern border overnight. "
        "You have thirty minutes to prepare your team before the next operation begins. Choose carefully.",
    ]
    _test_text_idx = 0
    _mixer_text_idx = 0

    def _voices_next_test_text(self):
        self._test_text_idx = (self._test_text_idx + 1) % len(self._TEST_TEXTS)
        self._voices_test_widget.delete("1.0", tk.END)
        self._voices_test_widget.insert("1.0", self._TEST_TEXTS[self._test_text_idx])
        total = len(self._TEST_TEXTS)
        self._voices_next_btn.config(
            text=f"{self._test_text_idx + 1}/{total}  Next Text")

    def _mixer_next_test_text(self):
        self._mixer_text_idx = (self._mixer_text_idx + 1) % len(self._TEST_TEXTS)
        self._preview_text_widget.delete("1.0", tk.END)
        self._preview_text_widget.insert("1.0", self._TEST_TEXTS[self._mixer_text_idx])
        total = len(self._TEST_TEXTS)
        self._mixer_next_btn.config(
            text=f"{self._mixer_text_idx + 1}/{total}\nNext Text")

    def _start_sapi_worker(self):
        """Start a persistent SAPI worker thread — COM initialised once, reused for every play."""
        import queue
        self._sapi_q = queue.Queue()
        self._sapi_stop_requested = False
        t = threading.Thread(target=self._sapi_worker, daemon=True)
        t.start()

    def _sapi_worker(self):
        import pythoncom, win32com.client, time
        pythoncom.CoInitialize()
        sapi = None
        try:
            sapi = win32com.client.Dispatch("SAPI.SpVoice")
            # Warmup in this apartment — runs once, silently
            sapi.Volume = 0
            sapi.Speak(" ", 0)   # synchronous, empty — forces audio device init
            sapi.Volume = 100
        except Exception:
            pass

        while True:
            try:
                item = self._sapi_q.get()
                if item is None:        # shutdown signal
                    break
                voice_key, text, on_done = item
                self._sapi_stop_requested = False
                if sapi:
                    try:
                        for token in sapi.GetVoices():
                            try:
                                if token.GetDescription() == voice_key:
                                    sapi.Voice = token
                                    break
                            except Exception:
                                pass
                        sapi.Speak(text, 1)   # SVSFlagsAsync — returns immediately
                        # Wait for SAPI to actually start speaking (loading a new voice can be slow)
                        deadline = time.time() + 3.0
                        while time.time() < deadline:
                            try:
                                if sapi.Status.RunningState == 2:
                                    break
                            except Exception:
                                break
                            time.sleep(0.02)
                        # Poll until finished or stop is requested
                        while True:
                            try:
                                if sapi.Status.RunningState != 2:   # 2 = SRSEIsSpeaking
                                    break
                            except Exception:
                                break
                            if self._sapi_stop_requested:
                                try:
                                    sapi.Speak("", 3)   # purge + async
                                except Exception:
                                    pass
                                break
                            time.sleep(0.05)
                    except Exception:
                        pass
                if on_done:
                    on_done()
            except Exception:
                pass

        pythoncom.CoUninitialize()

    def _voices_on_list_select(self, event=None):
        """Update the Selected Voice label; clear the other list so only one has a selection."""
        if getattr(self, '_selection_updating', False):
            return
        self._selection_updating = True
        try:
            # Use event.widget to determine which list was clicked — don't rely on which
            # list has a selection, because the other list may still hold a stale selection.
            clicked_left  = (event is not None and
                             hasattr(event, 'widget') and
                             event.widget is self._all_voices_list._tree)
            clicked_right = (event is not None and
                             hasattr(event, 'widget') and
                             event.widget is self._selected_voices_list._tree)

            entry = None
            if clicked_left:
                self._selected_voices_list.clear_selection()
                sel = self._all_voices_list.curselection()
                if sel:
                    entry = self._voices_available_data[sel[0]]
            elif clicked_right:
                self._all_voices_list.clear_selection()
                sel = self._selected_voices_list.curselection()
                if sel:
                    entry = self._selected_voices_data[sel[0]]

            if entry:
                label, _e, _k = entry
                display = label if len(label) <= 55 else label[:52] + "..."
                self._voices_selected_name_label.config(text=f"Selected Voice:  {display}")
            else:
                self._voices_selected_name_label.config(text="Selected Voice: —")
        finally:
            self._selection_updating = False

    def _voices_on_selected_click(self, event):
        """If the user clicks a SAPI voice that couldn't be located (greyed out, marked '(?)'),
        explain why and that re-installing the voice in Windows often fixes it."""
        try:
            iid = self._selected_voices_list.identify_row(event.y)
            if not iid:
                return
            idx = self._selected_voices_list.index_of_iid(iid)
            if idx is None or idx >= len(self._selected_voices_data):
                return
            label, eng, key = self._selected_voices_data[idx]
            fallback_nums = getattr(self, '_fallback_voice_numbers', {})
            if (eng, key) not in fallback_nums:
                return
            messagebox.showinfo(
                "Voice Unavailable",
                f"The voice \"{label}\" can't be found.\n\n"
                "Its installation files may have been moved or deleted. Reinstalling this "
                "SAPI online voice (Settings → Time & language → Speech → Manage voices — "
                "remove it, then add it again) should fix this.\n\n"
                "It's kept in this list, greyed out, so your selection isn't lost.")
        except Exception:
            pass

    # ── animation helpers ─────────────────────────────────────────────

    def _voices_start_anim(self, prefix, color, step=0):
        if not self.window.winfo_exists() or not self._voices_playing:
            return
        dots = (".", ". .", ". . .")
        self._voices_status_label.config(text=f"{prefix}{dots[step % 3]}", fg=color)
        self._voices_anim_id = self.window.after(
            400, self._voices_start_anim, prefix, color, step + 1)

    def _voices_stop_anim(self):
        aid = getattr(self, '_voices_anim_id', None)
        if aid:
            try:
                self.window.after_cancel(aid)
            except Exception:
                pass
        self._voices_anim_id = None

    def _voices_switch_to_playing(self, engine):
        """Called when synthesis finishes and audio playback begins."""
        self._voices_stop_anim()
        engine_names = {
            "piper":        "Piper",
            "piper_preset": "Piper",
            "kokoro":       "Kokoro",
            "kokoro_custom":"Kokoro",
            "sapi_preset":  "SAPI",
        }
        name  = engine_names.get(engine, engine.upper())
        color = self._ENGINE_COLORS.get(engine, "black")
        try:
            self._voices_status_label.config(text=f"{name}: Playing...", fg=color)
        except Exception:
            pass

    # ── play / stop ───────────────────────────────────────────────────

    def _voices_play(self):
        if self._voices_playing:
            # Stop — invalidate current token so any in-flight callback is ignored
            self._voices_play_token += 1
            self._sapi_stop_requested = True   # interrupt SAPI worker if it's speaking
            self.app.stop_speaking()
            self._voices_playing = False
            self._voices_stop_anim()
            self._voices_play_btn.config(text="▶ Play", bg="SystemButtonFace", fg="black")
            self._voices_status_label.config(text="", fg="black")
            return

        sel_left  = self._all_voices_list.curselection()
        sel_right = self._selected_voices_list.curselection()

        entry = None
        if sel_left:
            entry = self._voices_available_data[sel_left[0]]
        elif sel_right:
            entry = self._selected_voices_data[sel_right[0]]

        if not entry:
            messagebox.showwarning("No Voice Selected",
                                   "Select a voice from either list to test.",
                                   parent=self.window)
            return

        label, engine, key = entry
        text = self._voices_test_widget.get("1.0", tk.END).strip() or "Testing one two three."

        self._voices_playing = True
        self._voices_anim_id = None
        self._voices_play_token += 1
        token = self._voices_play_token
        self._voices_play_btn.config(text="⏹ Stop", bg="#cc3333", fg="white")

        color = self._ENGINE_COLORS.get(engine, "black")
        if engine == "sapi":
            self._voices_status_label.config(text="SAPI: Playing...", fg=color)
        else:
            engine_names = {"piper": "Piper", "piper_preset": "Piper",
                            "kokoro": "Kokoro", "kokoro_custom": "Kokoro",
                            "sapi_preset": "SAPI"}
            prefix = engine_names.get(engine, engine.upper())
            self._voices_start_anim(f"{prefix}: Generating", color)

        on_playing = lambda: self._voices_switch_to_playing(engine)

        def _done():
            if self.window.winfo_exists() and token == self._voices_play_token:
                self.window.after(0, self._voices_play_done)

        if engine == "sapi":
            self._sapi_q.put((key, text, _done))
        else:
            def _run():
                try:
                    if engine in ("piper", "piper_preset"):
                        model_key = f"preset:{key}" if engine == "piper_preset" else key
                        self.app._speak_with_piper(text, model_key, None,
                                                   _on_synth_done=on_playing)
                    elif engine in ("kokoro", "kokoro_custom"):
                        self.app._speak_with_kokoro(text, key,
                                                    _on_synth_done=on_playing)
                    elif engine == "sapi_preset":
                        self.app._speak_with_sapi_preset(text, key,
                                                         _on_synth_done=on_playing)
                except Exception:
                    pass
                finally:
                    _done()

            threading.Thread(target=_run, daemon=True).start()

    def _voices_play_done(self):
        self._voices_playing = False
        self._voices_stop_anim()
        try:
            self._voices_play_btn.config(text="▶ Play", bg="SystemButtonFace", fg="black")
            self._voices_status_label.config(text="", fg="black")
        except Exception:
            pass

    def _build_piper_tab(self, parent):
        # Scrollable: this tab stacks a lot of content vertically (lists row,
        # parameters, generating options) that can run past the bottom of
        # the window at high DPI scaling, with no other way to reach it.
        parent = self._make_scrollable(parent)

        # ── Top row: Available | arrows | Installed | Custom Voices ──
        lists = tk.Frame(parent)
        lists.pack(fill=tk.X, pady=(0, 6))

        left = tk.LabelFrame(lists, text="Available to Download", padx=5, pady=5)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))

        # Language filter inside the Available box
        filter_row = tk.Frame(left)
        filter_row.pack(fill=tk.X, pady=(0, 4))
        tk.Label(filter_row, text="Language:", font=("Helvetica", 8)).pack(side=tk.LEFT)
        self.lang_var = tk.StringVar(value="All")
        self.lang_combo = ttk.Combobox(filter_row, textvariable=self.lang_var,
                                        state="readonly", width=24)
        self.lang_combo.pack(side=tk.LEFT, padx=(4, 0))
        self.lang_combo.bind("<<ComboboxSelected>>", lambda _: self._refresh_lists())

        self._piper_not_installed_label = tk.Label(
            left, text="", font=("Helvetica", 8), fg="#aa4400", anchor="w", wraplength=200)
        self._piper_not_installed_label.pack(fill=tk.X)

        avail_inner = tk.Frame(left)
        avail_inner.pack(fill=tk.BOTH, expand=True)
        self.available_list = tk.Listbox(avail_inner, selectmode=tk.SINGLE,
                                          font=("Helvetica", 9), width=1, height=10)
        self.available_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_a = tk.Scrollbar(avail_inner, orient=tk.VERTICAL, command=self.available_list.yview)
        sb_a.pack(side=tk.RIGHT, fill=tk.Y)
        self.available_list.config(yscrollcommand=sb_a.set)
        self.available_list.bind("<<ListboxSelect>>", self._on_available_select)

        self.progress = ttk.Progressbar(left, orient=tk.HORIZONTAL,
                                        mode="determinate", length=1)
        self.progress.pack(fill=tk.X, pady=(4, 0))
        self.status_label = tk.Label(left, text="", font=("Helvetica", 8),
                                     anchor="w", fg="#444444")
        self.status_label.pack(fill=tk.X)

        mid = tk.Frame(lists)
        mid.pack(side=tk.LEFT, fill=tk.Y, padx=6)
        tk.Frame(mid).pack(fill=tk.BOTH, expand=True)
        self.download_btn = tk.Button(mid, text="Download >>", width=12,
                                      state=tk.DISABLED, command=self._start_download)
        self.download_btn.pack(pady=(0, 6))
        self.delete_btn = tk.Button(mid, text="Delete Voice", width=12,
                                    state=tk.DISABLED, command=self._delete_selected)
        self.delete_btn.pack()
        tk.Frame(mid).pack(fill=tk.BOTH, expand=True)

        right = tk.LabelFrame(lists, text="Installed Voices", padx=5, pady=5)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 4))
        self.installed_list = tk.Listbox(right, selectmode=tk.SINGLE,
                                          font=("Helvetica", 9), width=1, height=10)
        self.installed_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_i = tk.Scrollbar(right, orient=tk.VERTICAL, command=self.installed_list.yview)
        sb_i.pack(side=tk.RIGHT, fill=tk.Y)
        self.installed_list.config(yscrollcommand=sb_i.set)
        self.installed_list.bind("<<ListboxSelect>>", self._on_installed_select)

        cv_frame = tk.LabelFrame(lists, text="Custom Voices", padx=8, pady=6)
        cv_frame.pack(side=tk.LEFT, fill=tk.Y)
        cv_inner = tk.Frame(cv_frame)
        cv_inner.pack(fill=tk.BOTH, expand=True)
        self._piper_presets_list = tk.Listbox(cv_inner, selectmode=tk.SINGLE,
                                               font=("Helvetica", 10), width=16, height=10)
        self._piper_presets_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_p = tk.Scrollbar(cv_inner, orient=tk.VERTICAL,
                             command=self._piper_presets_list.yview)
        sb_p.pack(side=tk.RIGHT, fill=tk.Y)
        self._piper_presets_list.config(yscrollcommand=sb_p.set)
        cv_btn_row_p = tk.Frame(cv_frame)
        cv_btn_row_p.pack(fill=tk.X, pady=(4, 0))
        tk.Button(cv_btn_row_p, text="Load", width=8,
                  command=self._load_piper_preset).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(cv_btn_row_p, text="Delete", width=8,
                  command=self._delete_piper_preset).pack(side=tk.LEFT)

        # ── Middle row: Parameters (left) | Preview + Options (right) ──
        mid_pane = tk.Frame(parent)
        mid_pane.pack(fill=tk.X, pady=(0, 4))

        # ── Left: Piper Voice Parameters ─────────────────────────────
        params_frame = tk.LabelFrame(mid_pane, text="Piper Voice Parameters", padx=8, pady=6)
        params_frame.pack(side=tk.LEFT, anchor='nw', padx=(0, 6))

        sliders_col = params_frame

        self._piper_selected_voice_label = tk.Label(
            sliders_col,
            text="No voice selected  —  click an installed voice",
            font=("Helvetica", 9), fg="#777777", anchor='w', justify='left')
        self._piper_selected_voice_label.pack(anchor='w', pady=(0, 8))

        row_ns = tk.Frame(sliders_col)
        row_ns.pack(fill=tk.X, pady=(0, 2))
        tk.Label(row_ns, text="Noise Scale:", width=16, anchor='w').pack(side=tk.LEFT)
        tk.Scale(row_ns, variable=self.app.piper_noise_scale_var,
                 from_=0.0, to=2.0, resolution=0.001,
                 orient=tk.HORIZONTAL, length=180, showvalue=True,
                 digits=3).pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(row_ns, text="Expressiveness  (def. 0.667)",
                 font=("Helvetica", 8), fg="#888888").pack(side=tk.LEFT)

        row_nw = tk.Frame(sliders_col)
        row_nw.pack(fill=tk.X, pady=(0, 2))
        tk.Label(row_nw, text="Noise W:", width=16, anchor='w').pack(side=tk.LEFT)
        tk.Scale(row_nw, variable=self.app.piper_noise_w_var,
                 from_=0.0, to=2.0, resolution=0.001,
                 orient=tk.HORIZONTAL, length=180, showvalue=True,
                 digits=3).pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(row_nw, text="Duration variation  (def. 0.800)",
                 font=("Helvetica", 8), fg="#888888").pack(side=tk.LEFT)

        row_ss = tk.Frame(sliders_col)
        row_ss.pack(fill=tk.X, pady=(0, 2))
        tk.Label(row_ss, text="Sentence Silence:", width=16, anchor='w').pack(side=tk.LEFT)
        tk.Scale(row_ss, variable=self.app.piper_sentence_silence_var,
                 from_=0.0, to=5.0, resolution=0.1,
                 orient=tk.HORIZONTAL, length=180,
                 showvalue=True).pack(side=tk.LEFT, padx=(0, 6))
        tk.Label(row_ss, text="Silence between sentences  (def. 0.2)",
                 font=("Helvetica", 8), fg="#888888").pack(side=tk.LEFT)

        self._piper_pitch_selected_key    = None
        self._piper_pitch_var             = tk.DoubleVar(value=0.0)
        self._piper_pitch_preview_playing = False

        # Bottom section: pitch+reset on left, save section on right (bottom-aligned)
        bottom_frame = tk.Frame(sliders_col)
        bottom_frame.pack(fill=tk.X, pady=(0, 0))

        bottom_left = tk.Frame(bottom_frame)
        bottom_left.pack(side=tk.LEFT, anchor='nw')

        row_p = tk.Frame(bottom_left)
        row_p.pack(fill=tk.X, pady=(0, 2))
        tk.Label(row_p, text="Pitch:", width=16, anchor='w').pack(side=tk.LEFT)
        tk.Scale(row_p, variable=self._piper_pitch_var,
                 from_=-10, to=10, resolution=0.5,
                 orient=tk.HORIZONTAL, length=180,
                 showvalue=True).pack(side=tk.LEFT)

        reset_row = tk.Frame(bottom_left)
        reset_row.pack(fill=tk.X, pady=(4, 0))
        tk.Button(reset_row, text="Reset Sliders", width=14,
                  command=self._reset_piper_sliders).pack(side=tk.LEFT)

        # Save section — bottom of right column aligns with Reset Sliders
        save_right = tk.Frame(bottom_frame)
        save_right.pack(side=tk.LEFT, fill=tk.Y, padx=(12, 0))
        tk.Frame(save_right).pack(fill=tk.BOTH, expand=True)  # spacer
        tk.Label(save_right, text="Save voice as:", anchor='w',
                 font=("Helvetica", 9)).pack(anchor='w')
        self._piper_preset_name_var = tk.StringVar()
        tk.Entry(save_right, textvariable=self._piper_preset_name_var,
                 width=20).pack(fill=tk.X, pady=(2, 4))
        tk.Button(save_right, text="Save Custom Voice",
                  command=self._save_piper_preset).pack(anchor='w')

        # ── Right: Preview text + Generating options ──────────────────
        right_col = tk.Frame(mid_pane)
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, anchor='nw')

        preview_frame = tk.LabelFrame(right_col, text="Preview text", padx=8, pady=6)
        preview_frame.pack(fill=tk.X)

        preview_inner = tk.Frame(preview_frame)
        preview_inner.pack(fill=tk.BOTH, expand=True)
        self._piper_text_idx = 0
        self._piper_preview_text = tk.Text(preview_inner, height=3, width=1,
                                            wrap="word", font=("Helvetica", 9))
        self._piper_preview_text.insert("1.0", self._TEST_TEXTS[0])
        self._piper_preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        btn_col = tk.Frame(preview_inner)
        btn_col.pack(side=tk.LEFT, padx=(6, 0), fill=tk.Y)
        total_texts = len(self._TEST_TEXTS)
        self._piper_next_btn = tk.Button(btn_col,
                                          text=f"1/{total_texts}  Next Text",
                                          width=13,
                                          command=self._piper_next_test_text)
        self._piper_next_btn.pack(fill=tk.X)
        self._piper_play_btn = tk.Button(btn_col, text="▶ Preview",
                                          command=self._play_piper_test, width=13,
                                          pady=1)
        self._piper_play_btn.pack(fill=tk.X, pady=(4, 0))

        mode_frame = tk.LabelFrame(right_col, text="Generating options", padx=8, pady=6)
        mode_frame.pack(fill=tk.X, pady=(4, 0))

        tk.Label(mode_frame, text="When to start speaking", font=("Helvetica", 9, "bold"),
                 fg="#333333").pack(anchor="w", pady=(0, 2))
        tk.Radiobutton(mode_frame, text="Standard: synthesise full text first, then play",
                       variable=self.app.piper_streaming_var, value=0,
                       command=self._save_piper_mode).pack(anchor='w')
        tk.Radiobutton(mode_frame, text="Sentence-by-sentence: start speaking sooner",
                       variable=self.app.piper_streaming_var, value=1,
                       command=self._save_piper_mode).pack(anchor='w')

        tk.Label(mode_frame, text="How to split text", font=("Helvetica", 9, "bold"),
                 fg="#333333").pack(anchor="w", pady=(6, 2))
        tk.Radiobutton(mode_frame, text="Sentence boundaries only, splits at  .  !  ?",
                       variable=self.app.piper_split_mode_var, value=0,
                       command=self._save_piper_split_mode).pack(anchor='w')
        tk.Radiobutton(mode_frame, text="Smart split, also splits at commas in long sentences",
                       variable=self.app.piper_split_mode_var, value=1,
                       command=self._save_piper_split_mode).pack(anchor='w')

        self._refresh_piper_presets_list()
        self._update_piper_list_state()

    def _build_kokoro_tab(self, parent):
        # Scrollable in both directions: this tab packs several fixed-width
        # columns side by side (voice list, mixer, pitch, custom voices) that
        # can overflow sideways at high DPI scaling, on top of the usual
        # risk of overflowing past the bottom of the window.
        parent = self._make_scrollable(parent, horizontal=True)
        parent = tk.Frame(parent)
        parent.pack(fill=tk.BOTH, expand=True)


        voices_bin = os.path.join(APP_KOKORO_DIR, 'voices-v1.0.bin')
        voices_dict = load_voices_from_file(voices_bin) if os.path.exists(voices_bin) else KOKORO_VOICES
        self._kokoro_voice_options = [d for d, l in voices_dict.values()]
        self._kokoro_voice_ids     = list(voices_dict.keys())
        voice_options = self._kokoro_voice_options
        voice_ids     = self._kokoro_voice_ids

        # ── Two-column: voices list (left) | voice mixer (right) ──────
        top_pane = tk.Frame(parent)
        top_pane.pack(fill=tk.X, pady=(0, 4))

        # Left column — available voices
        left_col = tk.LabelFrame(top_pane,
                                  text="Available Voices (all active once model is downloaded)",
                                  padx=5, pady=5)
        left_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 4))

        vlist_inner = tk.Frame(left_col)
        vlist_inner.pack(fill=tk.BOTH, expand=True)
        self.kokoro_voice_list = tk.Listbox(vlist_inner, selectmode=tk.SINGLE,
                                             font=("Helvetica", 10), width=28, height=12)
        self.kokoro_voice_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_v = tk.Scrollbar(vlist_inner, orient=tk.VERTICAL, command=self.kokoro_voice_list.yview)
        sb_v.pack(side=tk.RIGHT, fill=tk.Y)
        self.kokoro_voice_list.config(yscrollcommand=sb_v.set)
        for voice_id, (display, lang) in voices_dict.items():
            self.kokoro_voice_list.insert(tk.END, display)
        self.kokoro_voice_list.bind('<<ListboxSelect>>', self._on_kokoro_voice_select)

        self._kokoro_not_installed_label = tk.Label(
            left_col, text="", font=("Helvetica", 9), fg="#aa4400",
            wraplength=180, justify="left", anchor="w")
        self._kokoro_not_installed_label.pack(anchor='w', pady=(4, 0))

        # mid_pane: mixer + pitch side by side at top, preview text spanning full width below
        mid_pane = tk.Frame(top_pane)
        mid_pane.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 4))

        upper_row = tk.Frame(mid_pane)
        upper_row.pack(fill=tk.X)

        mixer_frame = tk.LabelFrame(upper_row,
                                     text="Voice Mixer — blend two voices into a new one",
                                     padx=8, pady=6)
        mixer_frame.pack(side=tk.LEFT, anchor='nw', padx=(0, 4))

        self._mixer_a_var = tk.StringVar(value=voice_options[0] if voice_options else "")
        self._mixer_b_var = tk.StringVar(value=voice_options[1] if len(voice_options) > 1 else "")
        self._blend_var   = tk.IntVar(value=50)

        row_a = tk.Frame(mixer_frame)
        row_a.pack(fill=tk.X, pady=(0, 4))
        tk.Label(row_a, text="Voice A:", width=10, anchor='w').pack(side=tk.LEFT)
        ttk.Combobox(row_a, textvariable=self._mixer_a_var,
                     values=voice_options, state="readonly", width=22).pack(side=tk.LEFT)

        row_b = tk.Frame(mixer_frame)
        row_b.pack(fill=tk.X, pady=(0, 4))
        tk.Label(row_b, text="Voice B:", width=10, anchor='w').pack(side=tk.LEFT)
        ttk.Combobox(row_b, textvariable=self._mixer_b_var,
                     values=voice_options, state="readonly", width=22).pack(side=tk.LEFT)

        row_blend = tk.Frame(mixer_frame)
        row_blend.pack(fill=tk.X, pady=(0, 6))
        tk.Label(row_blend, text="Blend %:", width=10, anchor='w').pack(side=tk.LEFT)
        tk.Scale(row_blend, variable=self._blend_var, from_=0, to=100,
                 orient=tk.HORIZONTAL, length=160, showvalue=True).pack(side=tk.LEFT)

        _play_cmd = lambda: self._play_voice_preview(voice_ids, voice_options)
        self._preview_btn = tk.Button(mixer_frame, text="▶ Preview Mixed Voices",
                                       command=_play_cmd)
        self._preview_btn_original_command = _play_cmd
        self._kokoro_preview_playing = False
        self._preview_btn.pack(anchor='w', pady=(0, 6))

        tk.Label(mixer_frame, text="Save mixed voice as:", anchor='w',
                 font=("Helvetica", 9)).pack(anchor='w')
        self._mixer_name_var = tk.StringVar()
        tk.Entry(mixer_frame, textvariable=self._mixer_name_var,
                 width=22).pack(fill=tk.X, pady=(2, 4))
        tk.Button(mixer_frame, text="Save",
                  command=lambda: self._create_mixed_voice(voice_ids, voice_options),
                  width=10).pack(anchor='w')

        # Pitch Adjustment — right of mixer in upper_row, shrinks to content height
        self._pitch_selected_engine = None
        self._pitch_selected_key    = None
        self._pitch_preview_playing = False
        self._pitch_var             = tk.DoubleVar(value=0.0)

        pitch_frame = tk.LabelFrame(upper_row, text="Pitch Adjustment", padx=8, pady=6)
        pitch_frame.pack(side=tk.LEFT, anchor='nw')

        # Preview text — spans full mid_pane width below mixer+pitch
        preview_frame = tk.LabelFrame(mid_pane, text="Preview text", padx=8, pady=6)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 0))

        preview_inner = tk.Frame(preview_frame)
        preview_inner.pack(fill=tk.BOTH, expand=True)
        self._preview_text_widget = tk.Text(preview_inner, height=5, width=50,
                                             wrap="word", font=("Helvetica", 9))
        self._preview_text_widget.insert("1.0", self._TEST_TEXTS[0])
        self._preview_text_widget.pack(side=tk.LEFT, fill=tk.Y)

        self._mixer_text_idx = 0
        total_texts = len(self._TEST_TEXTS)
        self._mixer_next_btn = tk.Button(preview_inner,
                                          text=f"1/{total_texts}\nNext Text",
                                          width=10,
                                          command=self._mixer_next_test_text)
        self._mixer_next_btn.pack(side=tk.LEFT, padx=(6, 0), fill=tk.Y)

        self._pitch_voice_label = tk.Label(pitch_frame,
            text="Select an installed or\ncustom voice to modify.",
            font=("Helvetica", 9), fg="#777777", anchor='w', justify='left')
        self._pitch_voice_label.pack(anchor='w', pady=(0, 4))

        tk.Label(pitch_frame, text="Pitch (semitones):", anchor='w').pack(anchor='w')
        tk.Scale(pitch_frame, variable=self._pitch_var,
                 from_=-10, to=10, resolution=0.5,
                 orient=tk.HORIZONTAL, showvalue=True).pack(fill=tk.X, pady=(2, 4))

        pitch_btn_row = tk.Frame(pitch_frame)
        pitch_btn_row.pack(anchor='w', pady=(0, 6))
        tk.Button(pitch_btn_row, text="Reset", width=6,
                  command=lambda: self._pitch_var.set(0.0)).pack(side=tk.LEFT, padx=(0, 6))
        self._pitch_preview_btn = tk.Button(pitch_btn_row, text="▶ Preview pitch",
                                             width=14, command=self._pitch_preview)
        self._pitch_preview_btn.pack(side=tk.LEFT)

        tk.Label(pitch_frame, text="Save pitched voice as:", anchor='w').pack(anchor='w')
        self._pitch_name_var = tk.StringVar()
        tk.Entry(pitch_frame, textvariable=self._pitch_name_var,
                 width=32).pack(fill=tk.X, pady=(2, 6))
        tk.Button(pitch_frame, text="Save as Pitched",
                  command=self._save_pitched_voice).pack(anchor='w')

        # Custom Voices — fills remaining top_pane width
        cv_frame = tk.LabelFrame(top_pane, text="Custom Voices", padx=8, pady=6)
        cv_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, anchor='n')

        cv_inner = tk.Frame(cv_frame)
        cv_inner.pack(fill=tk.BOTH, expand=True)
        self.custom_voice_list = tk.Listbox(cv_inner, selectmode=tk.SINGLE,
                                             font=("Helvetica", 10), height=6)
        self.custom_voice_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_c = tk.Scrollbar(cv_inner, orient=tk.VERTICAL,
                             command=self.custom_voice_list.yview)
        sb_c.pack(side=tk.RIGHT, fill=tk.Y)
        self.custom_voice_list.config(yscrollcommand=sb_c.set)
        self.custom_voice_list.bind('<<ListboxSelect>>', self._on_custom_voice_select)
        cv_btn_row = tk.Frame(cv_frame)
        cv_btn_row.pack(fill=tk.X, pady=(4, 0))
        tk.Button(cv_btn_row, text="Load", width=8,
                  command=self._load_selected_voice).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(cv_btn_row, text="Delete", width=8,
                  command=self._delete_custom_voice).pack(side=tk.LEFT)

        # ── Bottom row: Import Voice Mixes (left) | Kokoro Settings (right) ──
        bottom_pane = tk.Frame(parent)
        bottom_pane.pack(fill=tk.X, pady=(0, 4))

        import_frame = tk.LabelFrame(bottom_pane, text="Import Voice Mixes", padx=6, pady=4)
        import_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 4), anchor='n')

        tk.Label(import_frame,
                 text=".npy files are premade Kokoro voice\nembeddings — find them online.",
                 font=("Helvetica", 8), fg="#555555", justify="left").pack(anchor='w', pady=(0, 6))

        self._import_path_var = tk.StringVar()
        tk.Button(import_frame, text="Browse .npy...", width=16,
                  command=self._browse_voice_file).pack(anchor='w')
        tk.Entry(import_frame, textvariable=self._import_path_var,
                 width=22, state="readonly").pack(fill=tk.X, pady=(2, 6))

        tk.Label(import_frame, text="Name:", anchor='w').pack(anchor='w')
        self._import_name_var = tk.StringVar()
        tk.Entry(import_frame, textvariable=self._import_name_var,
                 width=22).pack(fill=tk.X, pady=(2, 4))
        tk.Button(import_frame, text="Import", width=10,
                  command=self._import_voice_file).pack(anchor='w')

        # Kokoro Settings — fills remaining bottom_pane width
        ks_frame = tk.LabelFrame(bottom_pane, text="Kokoro Settings", padx=8, pady=4)
        ks_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, anchor='n')

        hw_row = tk.Frame(ks_frame)
        hw_row.pack(fill=tk.X, pady=(0, 4))
        tk.Label(hw_row, text="CPU Threads:", anchor='w').pack(side=tk.LEFT)
        threads_spinbox = tk.Spinbox(hw_row, from_=0, to=32, increment=1, width=4,
                                      textvariable=self.app.kokoro_num_threads_var,
                                      command=self._save_kokoro_threads)
        threads_spinbox.pack(side=tk.LEFT, padx=(4, 2))
        threads_spinbox.bind('<FocusOut>', lambda e: self._save_kokoro_threads())
        threads_spinbox.bind('<Return>', lambda e: self._save_kokoro_threads())
        tk.Label(hw_row, text="(0=auto)", font=("Helvetica", 8),
                 fg="#555555").pack(side=tk.LEFT, padx=(0, 16))

        _sep = tk.Frame(hw_row, width=1, bg="#cccccc")
        _sep.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 12))

        tk.Label(hw_row, text="GPU Provider:").pack(side=tk.LEFT)
        self._provider_combo = ttk.Combobox(
            hw_row,
            textvariable=self.app.kokoro_provider_var,
            values=["CPU", "DirectML", "CUDA"],
            state="readonly", width=12)
        self._provider_combo.pack(side=tk.LEFT, padx=(4, 6))
        self._provider_combo.bind("<<ComboboxSelected>>", lambda e: self._save_kokoro_provider())

        self._provider_info_label = tk.Label(hw_row, text="",
                 font=("Helvetica", 8), fg="#777777")
        self._provider_info_label.pack(side=tk.LEFT, padx=(4, 0))

        threading.Thread(target=self._detect_providers, daemon=True).start()

        tk.Label(ks_frame, text="Generating options",
                 font=("Helvetica", 9, "bold"), fg="#333333").pack(anchor="w", pady=(4, 2))

        speech_frame = tk.LabelFrame(ks_frame, text="When to start speaking", padx=8, pady=4)
        speech_frame.pack(fill=tk.X, pady=(0, 6))
        tk.Radiobutton(speech_frame,
                       text="Wait: synthesise all text first, then play  (smooth, no pauses between sentences)",
                       variable=self.app.kokoro_streaming_var, value=0,
                       command=self._save_kokoro_mode).pack(anchor='w')
        tk.Radiobutton(speech_frame,
                       text="Immediately: speak sentence by sentence as each one is ready  (fastest response, brief pauses between sentences)",
                       variable=self.app.kokoro_streaming_var, value=1,
                       command=self._save_kokoro_mode).pack(anchor='w')

        split_frame = tk.LabelFrame(ks_frame, text="How to split text", padx=8, pady=4)
        split_frame.pack(fill=tk.X)
        tk.Radiobutton(split_frame,
                       text="At sentence boundaries only, splits at  .  !  ?",
                       variable=self.app.kokoro_split_mode_var, value=0,
                       command=self._save_kokoro_split_mode).pack(anchor='w')
        tk.Radiobutton(split_frame,
                       text="Smart split, also splits at commas in long sentences  (shorter chunks, less waiting per chunk)",
                       variable=self.app.kokoro_split_mode_var, value=1,
                       command=self._save_kokoro_split_mode).pack(anchor='w')

        self._refresh_kokoro_status()
        self._refresh_custom_voices_list()
        self._update_kokoro_list_state()

    # -------------------------------------------------------- piper preset helpers

    _PIPER_LANG_MAP = {
        'en_US': 'English (USA)',    'en_GB': 'English (UK)',
        'en_AU': 'English (Australia)', 'de_DE': 'German',
        'fr_FR': 'French (France)',  'fr_BE': 'French (Belgium)',
        'es_ES': 'Spanish (Spain)',  'es_MX': 'Spanish (Mexico)',
        'it_IT': 'Italian',          'pt_PT': 'Portuguese (Portugal)',
        'pt_BR': 'Portuguese (Brazil)', 'nl_NL': 'Dutch',
        'nl_BE': 'Dutch (Belgium)',  'ru_RU': 'Russian',
        'zh_CN': 'Chinese (Simplified)', 'ja_JP': 'Japanese',
        'ko_KR': 'Korean',           'pl_PL': 'Polish',
        'cs_CZ': 'Czech',            'sk_SK': 'Slovak',
        'hu_HU': 'Hungarian',        'ro_RO': 'Romanian',
        'tr_TR': 'Turkish',          'da_DK': 'Danish',
        'fi_FI': 'Finnish',          'nb_NO': 'Norwegian',
        'sv_SE': 'Swedish',          'el_GR': 'Greek',
        'uk_UA': 'Ukrainian',        'ca_ES': 'Catalan',
        'is_IS': 'Icelandic',        'vi_VN': 'Vietnamese',
        'ar_JO': 'Arabic (Jordan)',  'ar_SA': 'Arabic (Saudi Arabia)',
        'fa_IR': 'Persian',          'he_IL': 'Hebrew',
        'hi_IN': 'Hindi',            'bn_BD': 'Bengali',
        'ta_IN': 'Tamil',            'te_IN': 'Telugu',
        'th_TH': 'Thai',             'id_ID': 'Indonesian',
        'ms_MY': 'Malay',            'tl_PH': 'Filipino',
        'sw_CD': 'Swahili',          'af_ZA': 'Afrikaans',
        'bg_BG': 'Bulgarian',        'hr_HR': 'Croatian',
        'sr_RS': 'Serbian',          'sl_SI': 'Slovenian',
        'lt_LT': 'Lithuanian',       'lv_LV': 'Latvian',
        'et_EE': 'Estonian',         'ur_PK': 'Urdu',
    }

    def _piper_key_label(self, key):
        """Convert 'en_US-amy-medium' to 'English (USA) - Amy - Medium'."""
        parts = key.split('-')
        if len(parts) >= 2:
            lang = self._PIPER_LANG_MAP.get(parts[0], parts[0].replace('_', ' '))
            name = ' - '.join(p.capitalize() for p in parts[1:])
            return f"{lang} - {name}"
        return key

    def _piper_next_test_text(self):
        self._piper_text_idx = (self._piper_text_idx + 1) % len(self._TEST_TEXTS)
        self._piper_preview_text.delete("1.0", tk.END)
        self._piper_preview_text.insert("1.0", self._TEST_TEXTS[self._piper_text_idx])
        total = len(self._TEST_TEXTS)
        self._piper_next_btn.config(text=f"{self._piper_text_idx + 1}/{total}  Next Text")

    def _reset_piper_sliders(self):
        """Reset all Piper synthesis sliders to their default values."""
        self.app.piper_noise_scale_var.set(0.667)
        self.app.piper_noise_w_var.set(0.800)
        self.app.piper_sentence_silence_var.set(0.2)
        self._piper_pitch_var.set(0.0)

    def _update_piper_selected_label(self, key):
        if not key:
            self._piper_selected_voice_label.config(
                text="No voice selected  —  click an installed voice", fg="#777777")
            return
        catalog = getattr(self, '_voice_catalog', {})
        info = catalog.get(key, {})
        name = self._format_voice_label(key, info) if info else key
        self._piper_selected_voice_label.config(text=f"Selected:  {name}", fg="#226600")

    def _refresh_piper_presets_list(self):
        """Refresh the saved presets listbox."""
        self._piper_presets_list.delete(0, tk.END)
        if os.path.isdir(APP_PIPER_PRESETS_DIR):
            for fname in sorted(os.listdir(APP_PIPER_PRESETS_DIR)):
                if fname.endswith('.json'):
                    self._piper_presets_list.insert(tk.END, f"{fname[:-5]} (custom)")

    def _play_piper_test(self):
        """Play the test text; button turns into Stop while playing."""
        model_key = self._piper_pitch_selected_key
        if not model_key:
            messagebox.showwarning("No Voice", "Click an installed voice to select it for preview.", parent=self.window)
            return
        widget = getattr(self, '_piper_preview_text', None)
        text = (widget.get("1.0", tk.END).strip() if widget else "") or "Hello, this is a test."

        self._piper_play_btn.config(text="⏹ Stop", command=self._stop_piper_test, bg="#cc3333", fg="white")

        pitch = self._piper_pitch_var.get() if not model_key.startswith('preset:') else None

        def _run():
            self.app._speak_with_piper(text, model_key, None, pitch_override=pitch)
            if self.window.winfo_exists():
                self.window.after(0, self._reset_piper_play_btn)

        threading.Thread(target=_run, daemon=True).start()

    def _stop_piper_test(self):
        """Cancel the currently playing test audio."""
        self.app.stop_speaking()
        self._reset_piper_play_btn()

    def _reset_piper_play_btn(self):
        """Restore the Preview button to its default state."""
        try:
            self._piper_play_btn.config(text="▶ Preview", command=self._play_piper_test,
                                         bg="SystemButtonFace", fg="black")
        except Exception:
            pass

    def _save_piper_preset(self):
        """Save current synthesis settings + selected voice as a named custom voice preset."""
        name = self._piper_preset_name_var.get().strip()
        if not name:
            messagebox.showwarning("No Name", "Enter a name for the preset.", parent=self.window)
            return
        model_key = self._piper_pitch_selected_key
        if not model_key:
            messagebox.showwarning("No Voice", "Click an installed voice to select it as the base.", parent=self.window)
            return
        if model_key.startswith('preset:'):
            messagebox.showwarning("Invalid Base",
                                   "Select an installed voice (not an existing preset) as the base.",
                                   parent=self.window)
            return
        os.makedirs(APP_PIPER_PRESETS_DIR, exist_ok=True)
        preset_path = os.path.join(APP_PIPER_PRESETS_DIR, f"{name}.json")
        if os.path.exists(preset_path):
            if not messagebox.askyesno("Overwrite?",
                                       f"A preset named '{name}' already exists. Overwrite it?",
                                       parent=self.window):
                return
        pitch = round(self._piper_pitch_var.get(), 1)
        data = {
            "base_voice": model_key,
            "noise_scale": round(self.app.piper_noise_scale_var.get(), 3),
            "noise_w": round(self.app.piper_noise_w_var.get(), 3),
            "sentence_silence": round(self.app.piper_sentence_silence_var.get(), 1),
        }
        if pitch != 0.0:
            data["type"] = "pitched"
            data["pitch"] = pitch
        with open(preset_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
        self._refresh_piper_presets_list()
        self._populate_voices_lists()
        self.app.refresh_voice_dropdowns()

    def _delete_piper_preset(self):
        """Delete the selected piper preset file."""
        sel = self._piper_presets_list.curselection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a preset to delete.", parent=self.window)
            return
        entry = self._piper_presets_list.get(sel[0])
        name = entry.split('  →  ')[0].strip().removesuffix(' (custom)')
        if not messagebox.askyesno("Delete Preset", f"Delete preset '{name}'?", parent=self.window):
            return
        preset_path = os.path.join(APP_PIPER_PRESETS_DIR, f"{name}.json")
        try:
            if os.path.exists(preset_path):
                os.remove(preset_path)
            self._refresh_piper_presets_list()
            self._populate_voices_lists()
            self.app.refresh_voice_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete preset:\n{e}", parent=self.window)

    def _load_piper_preset(self):
        sel = self._piper_presets_list.curselection()
        if not sel:
            return
        name = self._piper_presets_list.get(sel[0]).strip().removesuffix(' (custom)')
        preset_path = os.path.join(APP_PIPER_PRESETS_DIR, f"{name}.json")
        if not os.path.exists(preset_path):
            messagebox.showinfo("No Data", f"No data file found for '{name}'.",
                                parent=self.window)
            return
        try:
            with open(preset_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if 'noise_scale' in data:
                self.app.piper_noise_scale_var.set(data['noise_scale'])
            if 'noise_w' in data:
                self.app.piper_noise_w_var.set(data['noise_w'])
            if 'sentence_silence' in data:
                self.app.piper_sentence_silence_var.set(data['sentence_silence'])
            if data.get('type') == 'pitched':
                self._piper_pitch_var.set(data.get('pitch', 0.0))
            base_voice = data.get('base_voice', '')
            if base_voice:
                self._piper_pitch_selected_key = base_voice
                self._update_piper_selected_label(base_voice)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load preset:\n{e}", parent=self.window)

    # ──────────────────────────────────── Windows (SAPI) tab ──────────

    def _build_sapi_tab(self, parent):
        main_row = tk.Frame(parent)
        main_row.pack(fill=tk.BOTH, expand=True)
        # grid's weight-based distribution only kicks in for space left over
        # after each column's minimum (content-driven) width is satisfied -
        # columns 1/2 have several fixed character-width labels/entries that
        # can end up claiming most of the row at high DPI scaling, squeezing
        # column 0's voice list down to almost nothing. An explicit minsize
        # guarantees it stays usable regardless of what the other columns need.
        main_row.columnconfigure(0, weight=10, minsize=int(220 * get_dpi_scale()))
        main_row.columnconfigure(1, weight=3)
        main_row.columnconfigure(2, weight=2)
        main_row.rowconfigure(0, weight=1)

        # LEFT: SAPI Voices list
        left = tk.LabelFrame(main_row, text="SAPI Voices", padx=5, pady=5)
        left.grid(row=0, column=0, sticky='nsew', padx=(0, 3))
        sapi_inner = tk.Frame(left)
        sapi_inner.pack(fill=tk.BOTH, expand=True)
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)
        self._sapi_voices_list = tk.Listbox(sapi_inner, selectmode=tk.SINGLE,
                                             font=("Helvetica", 9), width=1, height=18)
        self._sapi_voices_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_sv = tk.Scrollbar(sapi_inner, orient=tk.VERTICAL,
                              command=self._sapi_voices_list.yview)
        sb_sv.pack(side=tk.RIGHT, fill=tk.Y)
        self._sapi_voices_list.config(yscrollcommand=sb_sv.set)
        self._sapi_voices_list.bind("<<ListboxSelect>>", self._on_sapi_voice_select)

        # Reuse the same enumeration as the Available Voices list — it falls back to
        # winreg when SAPI.GetVoices() is broken (e.g. blocked by NVDA, or corrupted by
        # Windows' "Online (Natural)" voices), so this list stays populated either way.
        self._sapi_voice_names = []
        for label, engine, key in self._enumerate_all_voices():
            if engine == "sapi":
                self._sapi_voice_names.append(key)
                self._sapi_voices_list.insert(tk.END, key)

        # CENTER: Parameters + Preview text stacked vertically
        center = tk.Frame(main_row)
        center.grid(row=0, column=1, sticky='nsew', padx=3)

        params_frame = tk.LabelFrame(center, text="Parameters", padx=8, pady=6)
        params_frame.pack(fill=tk.X, anchor='nw')

        self._sapi_pitch_selected_key = None
        self._sapi_pitch_var = tk.DoubleVar(value=0.0)

        self._sapi_selected_voice_label = tk.Label(
            params_frame,
            text="No voice selected  —  click a SAPI voice",
            font=("Helvetica", 9), fg="#777777", anchor='w', justify='left',
            width=42)
        self._sapi_selected_voice_label.pack(anchor='w', pady=(0, 8))

        row_p = tk.Frame(params_frame)
        row_p.pack(fill=tk.X, pady=(0, 2))
        tk.Label(row_p, text="Pitch:", width=10, anchor='w').pack(side=tk.LEFT)
        tk.Scale(row_p, variable=self._sapi_pitch_var,
                 from_=-10, to=10, resolution=0.5,
                 orient=tk.HORIZONTAL, length=1,
                 showvalue=True).pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_row = tk.Frame(params_frame)
        btn_row.pack(fill=tk.X, pady=(6, 0))
        tk.Button(btn_row, text="Reset Pitch", width=12,
                  command=lambda: self._sapi_pitch_var.set(0.0)).pack(side=tk.LEFT)

        save_frame = tk.LabelFrame(params_frame, text="Save as custom voice", padx=6, pady=4)
        save_frame.pack(fill=tk.X, pady=(10, 0))
        self._sapi_preset_name_var = tk.StringVar()
        tk.Entry(save_frame, textvariable=self._sapi_preset_name_var,
                 width=22).pack(fill=tk.X, pady=(0, 4))
        tk.Button(save_frame, text="Save Custom Voice",
                  command=self._save_sapi_preset).pack(anchor='w')

        preview_frame = tk.LabelFrame(center, text="Preview text", padx=8, pady=6)
        preview_frame.pack(fill=tk.X, pady=(8, 0))
        preview_inner = tk.Frame(preview_frame)
        preview_inner.pack(fill=tk.BOTH, expand=True)
        self._sapi_text_idx = 0
        self._sapi_preview_text = tk.Text(preview_inner, height=3, width=1,
                                           wrap="word", font=("Helvetica", 9))
        self._sapi_preview_text.insert("1.0", self._TEST_TEXTS[0])
        self._sapi_preview_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        btn_col = tk.Frame(preview_inner)
        btn_col.pack(side=tk.LEFT, padx=(6, 0), fill=tk.Y)
        total = len(self._TEST_TEXTS)
        self._sapi_next_btn = tk.Button(btn_col, text=f"1/{total}  Next Text",
                                         width=13, command=self._sapi_next_test_text)
        self._sapi_next_btn.pack(fill=tk.X)
        self._sapi_play_btn = tk.Button(btn_col, text="▶ Preview",
                                         command=self._play_sapi_test, width=13, pady=1)
        self._sapi_play_btn.pack(fill=tk.X, pady=(4, 0))

        # RIGHT: Custom Voices list
        cv_frame = tk.LabelFrame(main_row, text="Custom Voices", padx=8, pady=6)
        cv_frame.grid(row=0, column=2, sticky='nsew', padx=(3, 0))
        cv_frame.rowconfigure(0, weight=1)
        cv_frame.columnconfigure(0, weight=1)
        cv_inner = tk.Frame(cv_frame)
        cv_inner.pack(fill=tk.BOTH, expand=True)
        self._sapi_presets_list = tk.Listbox(cv_inner, selectmode=tk.SINGLE,
                                              font=("Helvetica", 10), width=1, height=18)
        self._sapi_presets_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_sp = tk.Scrollbar(cv_inner, orient=tk.VERTICAL,
                              command=self._sapi_presets_list.yview)
        sb_sp.pack(side=tk.RIGHT, fill=tk.Y)
        self._sapi_presets_list.config(yscrollcommand=sb_sp.set)
        cv_btn_row = tk.Frame(cv_frame)
        cv_btn_row.pack(fill=tk.X, pady=(4, 0))
        tk.Button(cv_btn_row, text="Load", width=8,
                  command=self._load_sapi_preset).pack(side=tk.LEFT, padx=(0, 4))
        tk.Button(cv_btn_row, text="Delete", width=8,
                  command=self._delete_sapi_preset).pack(side=tk.LEFT)

        self._refresh_sapi_presets_list()

    def _on_sapi_voice_select(self, event=None):
        sel = self._sapi_voices_list.curselection()
        if not sel:
            return
        name = self._sapi_voice_names[sel[0]]
        self._sapi_pitch_selected_key = name
        self._sapi_selected_voice_label.config(
            text=f"Selected:  {name}", fg="#226600")

    def _sapi_next_test_text(self):
        self._sapi_text_idx = (self._sapi_text_idx + 1) % len(self._TEST_TEXTS)
        self._sapi_preview_text.delete("1.0", tk.END)
        self._sapi_preview_text.insert("1.0", self._TEST_TEXTS[self._sapi_text_idx])
        total = len(self._TEST_TEXTS)
        self._sapi_next_btn.config(text=f"{self._sapi_text_idx + 1}/{total}  Next Text")

    def _play_sapi_test(self):
        voice_name = self._sapi_pitch_selected_key
        if not voice_name:
            messagebox.showwarning("No Voice",
                                   "Click a SAPI voice to select it for preview.",
                                   parent=self.window)
            return
        text = (self._sapi_preview_text.get("1.0", tk.END).strip()
                or "Hello, this is a test.")
        pitch = self._sapi_pitch_var.get()
        self._sapi_play_btn.config(text="⏹ Stop", command=self._stop_sapi_test,
                                    bg="#cc3333", fg="white")

        def _run():
            self.app._speak_with_sapi_pitched(text, voice_name, pitch)
            if self.window.winfo_exists():
                self.window.after(0, self._reset_sapi_play_btn)

        threading.Thread(target=_run, daemon=True).start()

    def _stop_sapi_test(self):
        self.app.stop_speaking()
        self._reset_sapi_play_btn()

    def _reset_sapi_play_btn(self):
        try:
            self._sapi_play_btn.config(text="▶ Preview", command=self._play_sapi_test,
                                        bg="SystemButtonFace", fg="black")
        except Exception:
            pass

    def _refresh_sapi_presets_list(self):
        self._sapi_presets_list.delete(0, tk.END)
        if os.path.isdir(APP_SAPI_PRESETS_DIR):
            for fname in sorted(os.listdir(APP_SAPI_PRESETS_DIR)):
                if fname.endswith('.json'):
                    self._sapi_presets_list.insert(tk.END, f"{fname[:-5]} (custom)")

    def _save_sapi_preset(self):
        name = self._sapi_preset_name_var.get().strip()
        if not name:
            messagebox.showwarning("No Name", "Enter a name for the custom voice.",
                                   parent=self.window)
            return
        voice_name = self._sapi_pitch_selected_key
        if not voice_name:
            messagebox.showwarning("No Voice",
                                   "Click a SAPI voice to select it as the base.",
                                   parent=self.window)
            return
        os.makedirs(APP_SAPI_PRESETS_DIR, exist_ok=True)
        preset_path = os.path.join(APP_SAPI_PRESETS_DIR, f"{name}.json")
        if os.path.exists(preset_path):
            if not messagebox.askyesno("Overwrite?",
                                       f"A custom voice named '{name}' already exists. Overwrite it?",
                                       parent=self.window):
                return
        data = {
            "base_voice": voice_name,
            "pitch": round(self._sapi_pitch_var.get(), 1),
        }
        with open(preset_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4)
        self._refresh_sapi_presets_list()
        self._populate_voices_lists()
        self.app.refresh_voice_dropdowns()

    def _delete_sapi_preset(self):
        sel = self._sapi_presets_list.curselection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a custom voice to delete.",
                                   parent=self.window)
            return
        name = self._sapi_presets_list.get(sel[0]).removesuffix(' (custom)')
        if not messagebox.askyesno("Delete", f"Delete '{name}'?", parent=self.window):
            return
        preset_path = os.path.join(APP_SAPI_PRESETS_DIR, f"{name}.json")
        try:
            if os.path.exists(preset_path):
                os.remove(preset_path)
            self._refresh_sapi_presets_list()
            self._populate_voices_lists()
            self.app.refresh_voice_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete:\n{e}", parent=self.window)

    def _load_sapi_preset(self):
        sel = self._sapi_presets_list.curselection()
        if not sel:
            return
        name = self._sapi_presets_list.get(sel[0]).removesuffix(' (custom)')
        preset_path = os.path.join(APP_SAPI_PRESETS_DIR, f"{name}.json")
        if not os.path.exists(preset_path):
            messagebox.showinfo("No Data", f"No file found for '{name}'.",
                                parent=self.window)
            return
        try:
            with open(preset_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            pitch = data.get('pitch', 0.0)
            self._sapi_pitch_var.set(pitch)
            voice_name = data.get('base_voice', '')
            if voice_name:
                self._sapi_pitch_selected_key = voice_name
                self._sapi_selected_voice_label.config(
                    text=f"Selected:  {voice_name}", fg="#226600")
                # Highlight the voice in the list
                for i, n in enumerate(self._sapi_voice_names):
                    if n == voice_name:
                        self._sapi_voices_list.selection_clear(0, tk.END)
                        self._sapi_voices_list.selection_set(i)
                        self._sapi_voices_list.see(i)
                        break
        except Exception as e:
            messagebox.showerror("Error", f"Could not load:\n{e}", parent=self.window)

    def _save_kokoro_threads(self):
        """Auto-save Kokoro CPU thread count to gametextreader_settings.json."""
        try:
            settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            settings['kokoro_num_threads'] = self.app.kokoro_num_threads_var.get()
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"[ERROR] Could not save Kokoro threads: {e}")

    def _detect_providers(self):
        """Detect available ONNX providers and update the dropdown + status label."""
        try:
            import onnxruntime as ort
            available = ort.get_available_providers()
        except Exception:
            available = []

        options = ["CPU"]
        if 'DmlExecutionProvider' in available:
            options.append("DirectML")
        if 'CUDAExecutionProvider' in available:
            options.append("CUDA")

        info_parts = []
        if 'DmlExecutionProvider' in available:
            info_parts.append("DirectML: GPU acceleration via DirectX 12 — works on any modern GPU.")
        if 'CUDAExecutionProvider' in available:
            info_parts.append("CUDA: NVIDIA GPU acceleration — fastest option for NVIDIA cards.")
        info_text = ("  ".join(info_parts) if info_parts
                     else "Only CPU is available on this system — no compatible GPU provider detected.")

        def _apply():
            if not self.window.winfo_exists():
                return
            self._provider_combo.config(values=options)
            # If saved selection isn't in available options, fall back to CPU
            if self.app.kokoro_provider_var.get() not in options:
                self.app.kokoro_provider_var.set("CPU")
                self._save_kokoro_provider()
            self._provider_info_label.config(text=info_text)

        self.window.after(0, _apply)

    def _save_kokoro_provider(self):
        """Auto-save GPU provider setting to gametextreader_settings.json."""
        try:
            settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            settings['kokoro_provider'] = self.app.kokoro_provider_var.get()
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            # Force model reload on next speak so the new provider takes effect
            if hasattr(self.app, '_kokoro_instance'):
                self.app._kokoro_instance = None
        except Exception as e:
            print(f"[ERROR] Could not save provider setting: {e}")

    def _save_piper_mode(self):
        """Auto-save Piper synthesis mode to gametextreader_settings.json."""
        try:
            settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            settings['piper_streaming'] = bool(self.app.piper_streaming_var.get())
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"[ERROR] Could not save Piper mode: {e}")

    def _save_kokoro_split_mode(self):
        try:
            settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            settings['kokoro_split_mode'] = self.app.kokoro_split_mode_var.get()
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"[ERROR] Could not save Kokoro split mode: {e}")

    def _save_piper_split_mode(self):
        try:
            settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            settings['piper_split_mode'] = self.app.piper_split_mode_var.get()
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"[ERROR] Could not save Piper split mode: {e}")

    def _save_kokoro_mode(self):
        """Auto-save Kokoro synthesis mode to gametextreader_settings.json."""
        try:
            settings = {}
            if os.path.exists(APP_SETTINGS_PATH):
                with open(APP_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            settings['kokoro_streaming'] = bool(self.app.kokoro_streaming_var.get())
            with open(APP_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"[ERROR] Could not save Kokoro mode: {e}")

    # ---------------------------------------------------------- engine helpers

    # -------------------------------------------------- update checking

    def _apply_cached_update_info(self):
        """Read update results cached by check_ai_voice_updates_on_startup() and update UI."""
        kokoro_info = getattr(self.app, '_kokoro_update_available', None)
        if kokoro_info:
            installed, latest = kokoro_info
            self.kokoro_status_label.config(
                text=f"Installed ({installed}) — update available: {latest}",
                fg="#885500")

    # ---------------------------------------------------------- engine helpers

    def _piper_exe_exists(self):
        # Piper engine is now bundled in the exe — always available
        return True

    def _refresh_engine_status(self):
        self.engine_status_label.config(text="Built-in", fg="green")

    def _open_engine_manage(self):
        webbrowser.open("https://github.com/OHF-Voice/piper1-gpl")

    # --------------------------------------------------------- Kokoro helpers

    def _kokoro_model_status(self):
        """Returns 'ok', 'corrupted', or 'not_installed'."""
        onnx_path = os.path.join(APP_KOKORO_DIR, 'kokoro-v1.0.int8.onnx')
        bin_path  = os.path.join(APP_KOKORO_DIR, 'voices-v1.0.bin')
        if not os.path.exists(onnx_path) or not os.path.exists(bin_path):
            return 'not_installed'
        try:
            if os.path.getsize(onnx_path) < 1_000_000:
                return 'corrupted'
            if os.path.getsize(bin_path) < 100_000:
                return 'corrupted'
            import numpy as np
            np.load(bin_path)
            return 'ok'
        except Exception:
            return 'corrupted'

    def _kokoro_model_exists(self):
        return self._kokoro_model_status() == 'ok'

    def _read_kokoro_version(self):
        try:
            with open(os.path.join(APP_KOKORO_DIR, 'kokoro_version.txt')) as f:
                return f.read().strip()
        except Exception:
            return None

    def _save_kokoro_version(self, version):
        try:
            with open(os.path.join(APP_KOKORO_DIR, 'kokoro_version.txt'), 'w') as f:
                f.write(version)
        except Exception:
            pass

    def _refresh_kokoro_status(self):
        status = self._kokoro_model_status()
        if status == 'ok':
            version = self._read_kokoro_version()
            ver_str = f"  ({version})" if version else ""
            self.kokoro_status_label.config(
                text=f"Installed{ver_str}", fg="green")
        elif status == 'corrupted':
            self.kokoro_status_label.config(
                text="Files corrupted or incomplete — re-download recommended", fg="red")
        else:
            self.kokoro_status_label.config(text="Not installed", fg="#999999")
        if hasattr(self, '_kokoro_not_installed_label'):
            self._update_kokoro_list_state()

    def _update_piper_list_state(self):
        """Gray out Piper available-voices list and disable download when engine is not installed."""
        installed = self._piper_exe_exists()
        if installed:
            self.available_list.config(state=tk.NORMAL, bg="white", fg="black")
            self.lang_combo.config(state="readonly")
            self._piper_not_installed_label.config(text="")
        else:
            self.available_list.config(state=tk.DISABLED, bg="#e8e8e8", fg="#aaaaaa")
            self.lang_combo.config(state=tk.DISABLED)
            self.download_btn.config(state=tk.DISABLED)
            self._piper_not_installed_label.config(
                text="Engine not installed — use Voices tab → Manage to install.",
                fg="#aa4400")

    def _update_kokoro_list_state(self):
        """Gray out Kokoro voice list when the model is not installed."""
        installed = self._kokoro_model_exists()
        if installed:
            self.kokoro_voice_list.config(state=tk.NORMAL, bg="white", fg="black")
            self._kokoro_not_installed_label.config(text="")
        else:
            self.kokoro_voice_list.config(state=tk.DISABLED, bg="#e8e8e8", fg="#aaaaaa")
            self._kokoro_not_installed_label.config(
                text="Kokoro model not installed — go to the Voices tab and click 'Manage' next to Kokoro to install it first.",
                fg="#aa4400")

    def _open_kokoro_manage(self):
        status_code = self._kokoro_model_status()
        kokoro_info = getattr(self.app, '_kokoro_update_available', None)
        if status_code == 'ok':
            version = self._read_kokoro_version()
            if kokoro_info:
                _, latest = kokoro_info
                status = f"Installed ({version}) — update available: {latest}"
                action = f"Download Update  ({latest})"
            else:
                status = f"Installed ({version})" if version else "Installed"
                action = "Re-download / Re-install"
        elif status_code == 'corrupted':
            status = "Files corrupted or incomplete"
            action = "Re-download to Fix"
        else:
            status = "Not installed"
            action = "Download and Install"
        icon_path = os.path.normpath(
            os.path.join(os.path.dirname(__file__), '..', '..', 'Assets', 'icon.ico'))
        can_delete = status_code in ('ok', 'corrupted')
        self._kokoro_popup = _ManageWindow(
            self.window, "Kokoro Model",
            ["Kokoro Text-to-Speech Model",
             "High-quality English voices (American & British).",
             "Available voices: 27"],
            "~116 MB  (model: ~88 MB + voices: ~28 MB)",
            status, action,
            on_download=self._start_kokoro_download,
            on_cancel=self._on_kokoro_download_cancelled,
            icon_path=icon_path,
            repo_url="https://github.com/thewh1teagle/kokoro-onnx",
            note_lines=["You will be notified by a popup if a model update is available.",
                        "Model files update independently of GameTextReader app updates.",
                        "Note: the Kokoro inference engine is bundled in the app."],
            delete_label="Delete Kokoro" if can_delete else None,
            on_delete=self._delete_kokoro_model if can_delete else None,
            status_color="green" if status_code == 'ok' else "#333333")
        if status_code == 'ok':
            self._kokoro_popup.set_progress(100)

    def _start_kokoro_download(self):
        threading.Thread(target=self._run_kokoro_download, daemon=True).start()

    def _on_kokoro_download_cancelled(self):
        for fname in (KOKORO_MODEL_FILE, KOKORO_VOICES_FILE):
            p = os.path.join(APP_KOKORO_DIR, fname)
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        self._refresh_kokoro_status()

    def _run_kokoro_download(self):
        popup = self._kokoro_popup
        try:
            import requests
            os.makedirs(APP_KOKORO_DIR, exist_ok=True)

            self.window.after(0, popup.set_status, "Checking latest Kokoro release...", "blue")
            api_resp = requests.get(KOKORO_GITHUB_API, timeout=15)
            api_resp.raise_for_status()
            release = api_resp.json()
            version = release.get('tag_name', 'unknown')

            assets = {a['name']: a['browser_download_url']
                      for a in release.get('assets', [])}

            wanted = [
                (KOKORO_VOICES_FILE, 28),
                (KOKORO_MODEL_FILE,  88),
            ]
            for filename, approx_mb in wanted:
                url = assets.get(filename)
                if not url:
                    raise RuntimeError(
                        f"'{filename}' not found in Kokoro release {version}.\n"
                        f"Available: {list(assets.keys())}")
                dest = os.path.join(APP_KOKORO_DIR, filename)
                self.window.after(0, popup.set_status,
                                  f"Downloading {filename} (~{approx_mb} MB)...", "blue")
                self.window.after(0, popup.set_progress, 0)
                with requests.get(url, stream=True, timeout=300) as r:
                    r.raise_for_status()
                    total = int(r.headers.get('content-length', approx_mb * 1_000_000))
                    downloaded = 0
                    with open(dest, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=65536):
                            if popup.cancelled:
                                raise RuntimeError("Cancelled")
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                pct = min(int(downloaded / total * 100), 100)
                                self.window.after(0, popup.set_progress, pct)

            self.window.after(0, self._on_kokoro_download_complete, True, None, version)
        except Exception as e:
            if not popup.cancelled:
                self.window.after(0, self._on_kokoro_download_complete, False, str(e), None)

    def _on_kokoro_download_complete(self, success, error, version=None):
        if success:
            if version:
                self._save_kokoro_version(version)
            ver_str = f"Installed ({version})" if version else "Installed"
            self._kokoro_popup.set_static_status(ver_str, "green")
            self._kokoro_popup.set_status("", "black")
            self._kokoro_popup.set_progress(100)
            self._kokoro_popup.mark_done()
            self._refresh_kokoro_status()
            self._populate_voices_lists()          # refresh Voice Dropdown tab
            if hasattr(self.app, "refresh_voice_dropdowns"):
                try:
                    self.app.refresh_voice_dropdowns()
                except Exception:
                    pass
        else:
            self._kokoro_popup.close()
            messagebox.showerror("Download Failed",
                                 f"Could not download Kokoro model.\n\n{error}",
                                 parent=self.window)

    def _delete_kokoro_model(self):
        if not messagebox.askyesno("Confirm Delete",
                                   "Delete the Kokoro model?\nAll Kokoro voices will be removed.",
                                   parent=self.window):
            return
        for fname in ('kokoro-v1.0.int8.onnx', 'voices-v1.0.bin'):
            p = os.path.join(APP_KOKORO_DIR, fname)
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        if hasattr(self.app, '_kokoro_instance'):
            self.app._kokoro_instance = None
        # Offer to delete custom voices too
        if os.path.isdir(APP_KOKORO_CUSTOM_DIR):
            custom_count = len([f for f in os.listdir(APP_KOKORO_CUSTOM_DIR)
                                 if f.endswith('.npy')])
            if custom_count > 0 and messagebox.askyesno(
                    "Delete Custom Voices?",
                    f"Also delete your {custom_count} custom voice(s)?",
                    parent=self.window):
                import shutil
                shutil.rmtree(APP_KOKORO_CUSTOM_DIR)
                self._refresh_custom_voices_list()
                self._populate_voices_lists()
        self._set_status("Kokoro model deleted.", "green")
        self._refresh_kokoro_status()
        self._populate_voices_lists()              # refresh Voice Dropdown tab
        if hasattr(self.app, "refresh_voice_dropdowns"):
            try:
                self.app.refresh_voice_dropdowns()
            except Exception:
                pass
        if hasattr(self, '_kokoro_popup'):
            try:
                self._kokoro_popup.close()
            except Exception:
                pass

    # -------------------------------------------------- custom voice helpers

    def _play_voice_preview(self, voice_ids, voice_options):
        """Synthesize the preview text with the current blend and play it."""
        if not self._kokoro_model_exists():
            messagebox.showwarning("Model Not Installed",
                                   "Download the Kokoro model first.",
                                   parent=self.window)
            return
        text = self._preview_text_widget.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("No Text", "Enter some preview text first.",
                                   parent=self.window)
            return
        self._preview_btn.config(state=tk.DISABLED, text="Generating.")
        self._preview_anim_id = None
        self._animate_generating(0)
        threading.Thread(
            target=self._run_preview,
            args=(text, voice_ids, voice_options),
            daemon=True
        ).start()

    def _run_preview(self, text, voice_ids, voice_options):
        wav_path = os.path.join(APP_KOKORO_DIR, '_preview.wav')
        try:
            import numpy as np
            import wave
            import ctypes
            import time

            a_label = self._mixer_a_var.get()
            b_label = self._mixer_b_var.get()
            blend   = self._blend_var.get()
            a_id = voice_ids[voice_options.index(a_label)] if a_label in voice_options else voice_ids[0]
            b_id = voice_ids[voice_options.index(b_label)] if b_label in voice_options else voice_ids[1]

            voices_path = os.path.join(APP_KOKORO_DIR, 'voices-v1.0.bin')
            voices_data = dict(np.load(voices_path))
            vec_a = voices_data[a_id]
            vec_b = voices_data[b_id]
            mixed = vec_a * (blend / 100.0) + vec_b * (1.0 - blend / 100.0)

            from kokoro_onnx import Kokoro as KokoroTTS
            model_path = os.path.join(APP_KOKORO_DIR, 'kokoro-v1.0.int8.onnx')
            if hasattr(self.app, '_kokoro_instance') and self.app._kokoro_instance:
                kokoro = self.app._kokoro_instance
            else:
                kokoro = KokoroTTS(model_path, voices_path)

            samples, sample_rate = kokoro.create(text, voice=mixed, speed=1.0, lang='en-us')

            pcm = (np.clip(samples, -1.0, 1.0) * 32767).astype(np.int16)
            with wave.open(wav_path, 'wb') as f:
                f.setnchannels(1)
                f.setsampwidth(2)
                f.setframerate(sample_rate)
                f.writeframes(pcm.tobytes())

            # Switch button to Stop before playing
            self._kokoro_preview_playing = True
            self.window.after(0, self._switch_preview_to_stop)

            # Play via MCI so we can stop mid-playback
            _winmm = ctypes.windll.winmm
            _alias = 'kokoroprev'
            _winmm.mciSendStringW(f'close {_alias}', None, 0, None)
            ret = _winmm.mciSendStringW(f'open "{wav_path}" type waveaudio alias {_alias}',
                                         None, 0, None)
            if ret == 0:
                _winmm.mciSendStringW(f'play {_alias}', None, 0, None)
                buf = ctypes.create_unicode_buffer(256)
                while self._kokoro_preview_playing:
                    _winmm.mciSendStringW(f'status {_alias} mode', buf, 256, None)
                    if buf.value in ('stopped', ''):
                        break
                    time.sleep(0.05)
                _winmm.mciSendStringW(f'stop {_alias}', None, 0, None)
                _winmm.mciSendStringW(f'close {_alias}', None, 0, None)
            else:
                import winsound
                winsound.PlaySound(wav_path, winsound.SND_FILENAME)

        except Exception as e:
            self.window.after(0, lambda: messagebox.showerror(
                "Preview Failed", f"Could not generate preview:\n{e}", parent=self.window))
        finally:
            self._kokoro_preview_playing = False
            try:
                if os.path.exists(wav_path):
                    os.remove(wav_path)
            except Exception:
                pass
            self.window.after(0, self._reset_preview_btn)

    def _switch_preview_to_stop(self):
        """Switch the preview button to a Stop button while audio is playing."""
        self._stop_generating_anim()
        self._preview_btn.config(state=tk.NORMAL, text="⏹ Stop",
                                 command=self._stop_kokoro_preview,
                                 bg="#cc3333", fg="white")

    def _stop_kokoro_preview(self):
        """Stop the currently playing Kokoro preview."""
        self._kokoro_preview_playing = False
        try:
            import ctypes
            ctypes.windll.winmm.mciSendStringW('stop kokoroprev', None, 0, None)
            ctypes.windll.winmm.mciSendStringW('close kokoroprev', None, 0, None)
        except Exception:
            pass

    def _animate_generating(self, step):
        """Cycle the preview button text while generation is in progress."""
        if not self.window.winfo_exists():
            return
        if self._preview_btn['state'] == tk.DISABLED:
            states = ["Generating.", "Generating. .", "Generating. . ."]
            self._preview_btn.config(text=states[step % 3])
            self._preview_anim_id = self.window.after(400, self._animate_generating, step + 1)

    def _stop_generating_anim(self):
        if hasattr(self, '_preview_anim_id') and self._preview_anim_id:
            try:
                self.window.after_cancel(self._preview_anim_id)
            except Exception:
                pass
            self._preview_anim_id = None

    def _reset_preview_btn(self):
        self._stop_generating_anim()
        self._preview_btn.config(state=tk.NORMAL, text="▶ Play Preview",
                                 command=self._preview_btn_original_command,
                                 bg="SystemButtonFace", fg="black")

    def _refresh_custom_voices_list(self):
        self.custom_voice_list.delete(0, tk.END)
        os.makedirs(APP_KOKORO_CUSTOM_DIR, exist_ok=True)
        for fname in sorted(os.listdir(APP_KOKORO_CUSTOM_DIR)):
            if fname.endswith('.npy'):
                self.custom_voice_list.insert(tk.END, fname[:-4])

    def _create_mixed_voice(self, voice_ids, voice_options):
        name = self._mixer_name_var.get().strip()
        if not name:
            messagebox.showwarning("Name Required",
                                   "Please enter a name for the custom voice.",
                                   parent=self.window)
            return
        if not self._kokoro_model_exists():
            messagebox.showwarning("Model Not Installed",
                                   "Download the Kokoro model first.",
                                   parent=self.window)
            return

        a_label = self._mixer_a_var.get()
        b_label = self._mixer_b_var.get()
        blend   = self._blend_var.get()  # % of voice A

        # Map display label back to voice_id
        a_id = voice_ids[voice_options.index(a_label)] if a_label in voice_options else voice_ids[0]
        b_id = voice_ids[voice_options.index(b_label)] if b_label in voice_options else voice_ids[1]

        os.makedirs(APP_KOKORO_CUSTOM_DIR, exist_ok=True)
        save_path = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{name}.npy")
        if os.path.exists(save_path):
            if not messagebox.askyesno("Name Already Exists",
                                       f"A custom voice named '{name}' already exists.\nOverwrite it?",
                                       parent=self.window):
                return

        try:
            import numpy as np
            voices_path = os.path.join(APP_KOKORO_DIR, 'voices-v1.0.bin')
            voices_data = dict(np.load(voices_path))
            vec_a = voices_data[a_id]
            vec_b = voices_data[b_id]
            mixed = vec_a * (blend / 100.0) + vec_b * (1.0 - blend / 100.0)

            os.makedirs(APP_KOKORO_CUSTOM_DIR, exist_ok=True)
            np.save(save_path, mixed)
            sidecar_path = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{name}.json")
            with open(sidecar_path, 'w', encoding='utf-8') as f:
                json.dump({"voice_a": a_label, "voice_b": b_label, "blend": blend}, f, indent=2)

            self._mixer_name_var.set("")
            self._refresh_custom_voices_list()
            self._populate_voices_lists()
            if hasattr(self.app, "refresh_voice_dropdowns"):
                self.app.refresh_voice_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Could not create voice:\n{e}", parent=self.window)

    def _browse_voice_file(self):
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Select Kokoro voice file",
            filetypes=[("Numpy voice file", "*.npy"), ("All files", "*.*")],
            parent=self.window)
        if path:
            self._import_path_var.set(path)
            # Pre-fill name from filename
            base = os.path.splitext(os.path.basename(path))[0]
            if not self._import_name_var.get():
                self._import_name_var.set(base)

    def _import_voice_file(self):
        src  = self._import_path_var.get().strip()
        name = self._import_name_var.get().strip()
        if not src or not os.path.exists(src):
            messagebox.showwarning("No File", "Please browse for a .npy voice file first.",
                                   parent=self.window)
            return
        if not name:
            messagebox.showwarning("Name Required", "Please enter a name for the voice.",
                                   parent=self.window)
            return
        os.makedirs(APP_KOKORO_CUSTOM_DIR, exist_ok=True)
        dest = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{name}.npy")
        if os.path.exists(dest):
            if not messagebox.askyesno("Name Already Exists",
                                       f"A custom voice named '{name}' already exists.\nOverwrite it?",
                                       parent=self.window):
                return
        try:
            import numpy as np
            import shutil
            # Validate it's a valid numpy file
            np.load(src)
            shutil.copy2(src, dest)

            self._import_path_var.set("")
            self._import_name_var.set("")
            self._refresh_custom_voices_list()
            self._populate_voices_lists()
            if hasattr(self.app, "refresh_voice_dropdowns"):
                self.app.refresh_voice_dropdowns()
        except Exception as e:
            messagebox.showerror("Import Failed", f"Could not import voice file:\n{e}",
                                 parent=self.window)

    def _on_kokoro_voice_select(self, _event=None):
        sel = self.kokoro_voice_list.curselection()
        if not sel:
            return
        self._pitch_selected_engine = 'builtin'
        self._pitch_selected_key    = self._kokoro_voice_ids[sel[0]]
        display = self._kokoro_voice_options[sel[0]]
        self._pitch_voice_label.config(
            text=f"Selected:  {display}  (built-in)", fg="#226600")
        self.custom_voice_list.selection_clear(0, tk.END)
        self._pitch_var.set(0.0)

    def _on_custom_voice_select(self, _event=None):
        sel = self.custom_voice_list.curselection()
        if not sel:
            return
        name = self.custom_voice_list.get(sel[0]).strip()
        self._pitch_selected_engine = 'custom'
        self._pitch_selected_key    = name
        self._pitch_voice_label.config(
            text=f"Selected:  {name}  (custom)", fg="#7700aa")
        self.kokoro_voice_list.selection_clear(0, tk.END)
        self._pitch_var.set(0.0)

    def _apply_pitch_shift(self, samples, sample_rate, semitones):
        if semitones == 0.0:
            return samples
        try:
            from pedalboard import PitchShift as _PitchShift
            import numpy as np
            return _PitchShift(semitones=float(semitones))(
                samples.astype(np.float32), sample_rate=int(sample_rate))
        except Exception:
            return samples

    def _pitch_preview(self):
        if not self._kokoro_model_exists():
            messagebox.showwarning("Model Not Installed",
                                   "Download the Kokoro model first.",
                                   parent=self.window)
            return
        if not self._pitch_selected_engine:
            messagebox.showwarning("No Voice Selected",
                                   "Select a voice from either list above.",
                                   parent=self.window)
            return
        semitones = self._pitch_var.get()
        text = self._preview_text_widget.get("1.0", tk.END).strip() \
               or "Hello, this is a pitch preview."
        self._pitch_preview_btn.config(state=tk.DISABLED, text="Generating.")
        self._pitch_anim_id = None
        self._animate_pitch_generating(0)
        self._pitch_preview_playing = True
        threading.Thread(
            target=self._run_pitch_preview,
            args=(text, self._pitch_selected_engine,
                  self._pitch_selected_key, semitones),
            daemon=True).start()

    def _run_pitch_preview(self, text, engine, key, semitones):
        import ctypes, time
        wav_path = os.path.join(APP_KOKORO_DIR, '_pitch_preview.wav')
        try:
            import numpy as np
            import wave
            from kokoro_onnx import Kokoro as KokoroTTS
            model_path  = os.path.join(APP_KOKORO_DIR, 'kokoro-v1.0.int8.onnx')
            voices_path = os.path.join(APP_KOKORO_DIR, 'voices-v1.0.bin')
            if hasattr(self.app, '_kokoro_instance') and self.app._kokoro_instance:
                kokoro = self.app._kokoro_instance
            else:
                kokoro = KokoroTTS(model_path, voices_path)
            if engine == 'builtin':
                voice_arg = key
            else:
                custom_path = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{key}.npy")
                voice_arg = np.load(custom_path)
            samples, sample_rate = kokoro.create(text, voice=voice_arg,
                                                  speed=1.0, lang='en-us')
            samples = self._apply_pitch_shift(samples, sample_rate, semitones)
            pcm = (np.clip(samples, -1.0, 1.0) * 32767).astype(np.int16)
            with wave.open(wav_path, 'wb') as f:
                f.setnchannels(1); f.setsampwidth(2)
                f.setframerate(sample_rate); f.writeframes(pcm.tobytes())

            self.window.after(0, self._switch_pitch_to_stop)
            _winmm = ctypes.windll.winmm
            _alias = 'pitchprev'
            _winmm.mciSendStringW(f'close {_alias}', None, 0, None)
            ret = _winmm.mciSendStringW(
                f'open "{wav_path}" type waveaudio alias {_alias}', None, 0, None)
            if ret == 0:
                _winmm.mciSendStringW(f'play {_alias}', None, 0, None)
                buf = ctypes.create_unicode_buffer(256)
                while self._pitch_preview_playing:
                    _winmm.mciSendStringW(f'status {_alias} mode', buf, 256, None)
                    if buf.value in ('stopped', ''):
                        break
                    time.sleep(0.05)
                _winmm.mciSendStringW(f'stop {_alias}', None, 0, None)
                _winmm.mciSendStringW(f'close {_alias}', None, 0, None)
            else:
                import winsound
                winsound.PlaySound(wav_path, winsound.SND_FILENAME)
        except Exception as e:
            self.window.after(0, lambda: messagebox.showerror(
                "Preview Failed", f"Could not generate pitch preview:\n{e}",
                parent=self.window))
        finally:
            self._pitch_preview_playing = False
            try:
                if os.path.exists(wav_path):
                    os.remove(wav_path)
            except Exception:
                pass
            self.window.after(0, self._reset_pitch_preview_btn)

    def _switch_pitch_to_stop(self):
        try:
            if not self.window.winfo_exists():
                return
            self._stop_pitch_generating_anim()
            self._pitch_preview_btn.config(
                state=tk.NORMAL, text="⏹ Stop",
                command=self._stop_pitch_preview,
                bg="#cc3333", fg="white")
        except Exception:
            pass

    def _stop_pitch_preview(self):
        self._pitch_preview_playing = False
        try:
            import ctypes
            ctypes.windll.winmm.mciSendStringW('stop pitchprev', None, 0, None)
            ctypes.windll.winmm.mciSendStringW('close pitchprev', None, 0, None)
        except Exception:
            pass

    def _animate_pitch_generating(self, step):
        if not self.window.winfo_exists():
            return
        if self._pitch_preview_btn['state'] == tk.DISABLED:
            states = ["Generating.", "Generating. .", "Generating. . ."]
            self._pitch_preview_btn.config(text=states[step % 3])
            self._pitch_anim_id = self.window.after(400, self._animate_pitch_generating, step + 1)

    def _stop_pitch_generating_anim(self):
        if hasattr(self, '_pitch_anim_id') and self._pitch_anim_id:
            try:
                self.window.after_cancel(self._pitch_anim_id)
            except Exception:
                pass
            self._pitch_anim_id = None

    def _reset_pitch_preview_btn(self):
        self._stop_pitch_generating_anim()
        try:
            self._pitch_preview_btn.config(
                state=tk.NORMAL, text="▶ Preview pitch",
                command=self._pitch_preview,
                bg="SystemButtonFace", fg="black")
        except Exception:
            pass

    def _save_pitched_voice(self):
        if not self._pitch_selected_engine:
            messagebox.showwarning("No Voice Selected",
                                   "Select a voice from either list above.",
                                   parent=self.window)
            return
        name = self._pitch_name_var.get().strip()
        if not name:
            messagebox.showwarning("Name Required",
                                   "Enter a name in the 'Save voice as' field below Preview.",
                                   parent=self.window)
            return
        full_name  = f"{name} (pitched)"
        save_path  = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{full_name}.npy")
        sidecar    = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{full_name}.json")
        if os.path.exists(save_path):
            if not messagebox.askyesno("Already Exists",
                                       f"'{full_name}' already exists. Overwrite?",
                                       parent=self.window):
                return
        try:
            import numpy as np
            voices_path = os.path.join(APP_KOKORO_DIR, 'voices-v1.0.bin')
            if self._pitch_selected_engine == 'builtin':
                voices_data = dict(np.load(voices_path))
                embedding = voices_data[self._pitch_selected_key]
            else:
                custom_path = os.path.join(APP_KOKORO_CUSTOM_DIR,
                                            f"{self._pitch_selected_key}.npy")
                embedding = np.load(custom_path)
            os.makedirs(APP_KOKORO_CUSTOM_DIR, exist_ok=True)
            np.save(save_path, embedding)
            with open(sidecar, 'w', encoding='utf-8') as f:
                json.dump({
                    "type":        "pitched",
                    "base_voice":  self._pitch_selected_key,
                    "base_engine": self._pitch_selected_engine,
                    "pitch":       self._pitch_var.get(),
                }, f, indent=2)
            self._pitch_name_var.set("")
            self._refresh_custom_voices_list()
            self._populate_voices_lists()
            if hasattr(self.app, "refresh_voice_dropdowns"):
                self.app.refresh_voice_dropdowns()
        except Exception as e:
            messagebox.showerror("Error", f"Could not save pitched voice:\n{e}",
                                 parent=self.window)

    def _load_selected_voice(self):
        sel = self.custom_voice_list.curselection()
        if not sel:
            return
        name = self.custom_voice_list.get(sel[0]).strip()
        sidecar = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{name}.json")
        if not os.path.exists(sidecar):
            messagebox.showinfo(
                "No Mix Info",
                f"'{name}' has no saved mix info.\n\n"
                "Load is only available for voices created with the mixer here.\n"
                "Imported voices do not carry Voice A / B / Blend data.",
                parent=self.window)
            return
        try:
            with open(sidecar, 'r', encoding='utf-8') as f:
                data = json.load(f)
            voice_type = data.get('type', 'mixed')
            if voice_type == 'pitched':
                self._pitch_var.set(data.get('pitch', 0.0))
                base = data.get('base_voice', '')
                self._pitch_selected_engine = data.get('base_engine', 'builtin')
                self._pitch_selected_key    = base
                self._pitch_voice_label.config(
                    text=f"Selected:  {base}  ({self._pitch_selected_engine})",
                    fg="#226600" if self._pitch_selected_engine == 'builtin' else "#7700aa")
                self._pitch_name_var.set(name.replace(' (pitched)', ''))
            else:
                self._mixer_a_var.set(data.get('voice_a', ''))
                self._mixer_b_var.set(data.get('voice_b', ''))
                self._blend_var.set(data.get('blend', 50))
                self._mixer_name_var.set(name)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load mix info:\n{e}", parent=self.window)

    def _delete_custom_voice(self):
        sel = self.custom_voice_list.curselection()
        if not sel:
            return
        name = self.custom_voice_list.get(sel[0]).strip()
        if not messagebox.askyesno("Confirm Delete", f"Delete custom voice:\n{name}?",
                                   parent=self.window):
            return
        path = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{name}.npy")
        sidecar = os.path.join(APP_KOKORO_CUSTOM_DIR, f"{name}.json")
        try:
            if os.path.exists(path):
                os.remove(path)
            if os.path.exists(sidecar):
                os.remove(sidecar)
        except Exception as e:
            messagebox.showerror("Error", f"Could not delete:\n{e}", parent=self.window)
            return
        self._refresh_custom_voices_list()
        self._populate_voices_lists()
        if hasattr(self.app, "refresh_voice_dropdowns"):
            self.app.refresh_voice_dropdowns()

    # --------------------------------------------------------- voice catalog

    _CATALOG_CACHE_PATH = os.path.join(APP_PIPER_DIR, '_voices_catalog_cache.json')
    _CATALOG_MAX_AGE_DAYS = 7

    def _fetch_voice_catalog(self):
        threading.Thread(target=self._run_fetch_catalog, daemon=True).start()

    def _run_fetch_catalog(self):
        import time
        cache = self._CATALOG_CACHE_PATH
        max_age = self._CATALOG_MAX_AGE_DAYS * 86400
        # Use cached file if it exists and is fresh enough
        if os.path.exists(cache):
            age = time.time() - os.path.getmtime(cache)
            if age < max_age:
                try:
                    with open(cache, 'r', encoding='utf-8') as f:
                        catalog = json.load(f)
                    self.window.after(0, self._on_catalog_loaded, catalog)
                    return
                except Exception:
                    pass  # cache corrupt — fall through to network fetch
        try:
            import requests
            r = requests.get(PIPER_VOICES_JSON_URL, timeout=20)
            r.raise_for_status()
            catalog = r.json()
            try:
                os.makedirs(APP_PIPER_DIR, exist_ok=True)
                with open(cache, 'w', encoding='utf-8') as f:
                    json.dump(catalog, f)
            except Exception:
                pass
            self.window.after(0, self._on_catalog_loaded, catalog)
        except Exception as e:
            # Network failed — try stale cache as fallback
            if os.path.exists(cache):
                try:
                    with open(cache, 'r', encoding='utf-8') as f:
                        catalog = json.load(f)
                    self.window.after(0, self._on_catalog_loaded, catalog)
                    return
                except Exception:
                    pass
            self.window.after(0, self._on_catalog_error, str(e))

    def _on_catalog_loaded(self, catalog):
        self._voice_catalog = catalog

        # Build language list for filter
        langs = set()
        for info in catalog.values():
            lang = info.get('language', {}).get('name_english', '')
            region = info.get('language', {}).get('country_english', '')
            if lang:
                langs.add(f"{lang} ({region})" if region else lang)
        lang_list = ["All"] + sorted(langs)
        self.lang_combo["values"] = lang_list

        self._set_status("Ready", "black")
        self._refresh_lists()

    def _on_catalog_error(self, error):
        self._set_status("Could not load voice list (offline?)", "red")

    def _installed_keys(self):
        keys = set()
        if os.path.isdir(APP_PIPER_VOICES_DIR):
            for f in os.listdir(APP_PIPER_VOICES_DIR):
                if f.endswith('.onnx'):
                    keys.add(f[:-5])
        return keys

    def _refresh_lists(self):
        installed = self._installed_keys()
        lang_filter = self.lang_var.get()

        def _lang_sort_key(item):
            k, inf = item
            lg = inf.get('language', {})
            return (lg.get('name_english', 'zzz'), lg.get('country_english', ''),
                    inf.get('name', k))

        # Available list — sorted by language then voice name
        self.available_list.delete(0, tk.END)
        self._available_map.clear()
        idx = 0
        for key, info in sorted(self._voice_catalog.items(), key=_lang_sort_key):
            if key in installed:
                continue
            lang = info.get('language', {})
            lang_name = lang.get('name_english', '')
            region = lang.get('country_english', '')
            lang_label = f"{lang_name} ({region})" if region else lang_name
            if lang_filter != "All" and lang_label != lang_filter:
                continue
            label = self._format_voice_label(key, info)
            self.available_list.insert(tk.END, label)
            self._available_map[idx] = key
            idx += 1

        # Installed list — sorted by language then voice name
        self.installed_list.delete(0, tk.END)
        self._installed_map.clear()
        idx = 0
        installed_items = [(k, self._voice_catalog.get(k, {})) for k in installed]
        for key, info in sorted(installed_items, key=_lang_sort_key):
            label = self._format_voice_label(key, info) if info else key
            self.installed_list.insert(tk.END, label)
            self._installed_map[idx] = key
            idx += 1

        self.download_btn.config(state=tk.DISABLED)
        self.delete_btn.config(state=tk.DISABLED)

    def _format_voice_label(self, key, info):
        """Build a clean human-readable label — language first for easy scanning."""
        lang = info.get('language', {})
        name = info.get('name', key.split('-')[1] if '-' in key else key).capitalize()
        lang_name = lang.get('name_english', '')
        region = lang.get('country_english', '')
        lang_str = f"{lang_name} ({region})" if region else lang_name
        quality = info.get('quality', '').capitalize()
        num_speakers = info.get('num_speakers', 1)
        speakers_str = f"  · {num_speakers} spk" if num_speakers > 1 else ""
        size_mb = self._get_size_mb(info)
        size_str = f"  · ~{size_mb} MB" if size_mb else ""

        return f"{lang_str}  ·  {name}  ·  {quality}{speakers_str}{size_str}"

    def _get_size_mb(self, info):
        try:
            for path, meta in info.get('files', {}).items():
                if path.endswith('.onnx'):
                    return round(meta['size_bytes'] / 1_000_000)
        except Exception:
            pass
        return None

    def _get_file_urls(self, key):
        """Return list of (url, filename) tuples for a voice key."""
        info = self._voice_catalog.get(key, {})
        urls = []
        for file_path in info.get('files', {}).keys():
            url = f"{PIPER_HF_BASE}/{file_path}"
            filename = os.path.basename(file_path)
            urls.append((url, filename))
        return urls

    # ---------------------------------------------------------- list selection

    def _on_available_select(self, _event=None):
        sel = self.available_list.curselection()
        if not sel:
            self.download_btn.config(state=tk.DISABLED)
            return
        if self._piper_exe_exists():
            self.download_btn.config(state=tk.NORMAL)
        self.installed_list.selection_clear(0, tk.END)
        self.delete_btn.config(state=tk.DISABLED)

    def _on_installed_select(self, _event=None):
        sel = self.installed_list.curselection()
        if not sel:
            self.delete_btn.config(state=tk.DISABLED)
            return
        key = self._installed_map.get(sel[0], "")
        self.delete_btn.config(state=tk.NORMAL)
        self.available_list.selection_clear(0, tk.END)
        self.download_btn.config(state=tk.DISABLED)
        # Selecting from Installed Voices also picks the voice for Parameters
        if key:
            self._piper_pitch_selected_key = key
            self._update_piper_selected_label(key)

    # ---------------------------------------------------------------- download

    def _start_download(self):
        if not self._piper_exe_exists():
            messagebox.showinfo(
                "Piper Engine Not Installed",
                "The Piper engine is not installed.\n\n"
                "Go to the Voices tab and click 'Manage' next to 'AI Voice Piper' to install it first.",
                parent=self.window)
            return
        sel = self.available_list.curselection()
        if not sel:
            return
        key = self._available_map.get(sel[0])
        if not key:
            return
        self.download_btn.config(state=tk.DISABLED)
        self.delete_btn.config(state=tk.DISABLED)
        self._set_status(f"Downloading {key}...", "blue")
        self.progress["value"] = 0
        threading.Thread(target=self._run_download, args=(key,), daemon=True).start()

    def _run_download(self, key):
        try:
            import requests
            urls = self._get_file_urls(key)
            if not urls:
                raise RuntimeError("No file URLs found for this voice.")

            total_files = len(urls)
            for i, (url, filename) in enumerate(urls):
                dest = os.path.join(APP_PIPER_VOICES_DIR, filename)
                self.window.after(0, self._set_status,
                                  f"Downloading {filename} ({i+1}/{total_files})...", "blue")
                with requests.get(url, stream=True, timeout=120) as r:
                    r.raise_for_status()
                    file_total = int(r.headers.get('content-length', 0))
                    downloaded = 0
                    with open(dest, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=65536):
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                if file_total:
                                    # Overall progress across all files
                                    pct = int((i / total_files + downloaded / file_total / total_files) * 100)
                                    self.window.after(0, self._set_progress, min(pct, 99))

            self.window.after(0, self._on_download_complete, key, True, None)
        except Exception as e:
            # Clean up partial files
            for _, filename in self._get_file_urls(key):
                p = os.path.join(APP_PIPER_VOICES_DIR, filename)
                try:
                    if os.path.exists(p):
                        os.remove(p)
                except Exception:
                    pass
            self.window.after(0, self._on_download_complete, key, False, str(e))

    def _set_progress(self, pct):
        self.progress["value"] = pct

    def _on_download_complete(self, key, success, error):
        self.progress["value"] = 0
        if success:
            self._set_status(f"Installed: {key}", "green")
            self._refresh_lists()
            self._populate_voices_lists()          # refresh Voice Dropdown tab
            if hasattr(self.app, "refresh_voice_dropdowns"):
                try:
                    self.app.refresh_voice_dropdowns()
                except Exception:
                    pass
        else:
            self._set_status("Download failed.", "red")
            messagebox.showerror("Download Failed",
                                 f"Could not download voice.\n\n{error}",
                                 parent=self.window)

    # ------------------------------------------------------------------ delete

    def _delete_selected(self):
        sel = self.installed_list.curselection()
        if not sel:
            return
        key = self._installed_map.get(sel[0])
        if not key:
            return

        # Find any custom presets that use this voice as their base
        bound_presets = []
        if os.path.isdir(APP_PIPER_PRESETS_DIR):
            for fname in sorted(os.listdir(APP_PIPER_PRESETS_DIR)):
                if fname.endswith('.json'):
                    try:
                        with open(os.path.join(APP_PIPER_PRESETS_DIR, fname),
                                  'r', encoding='utf-8') as _f:
                            data = json.load(_f)
                        if data.get('base_voice') == key:
                            bound_presets.append(fname[:-5])
                    except Exception:
                        pass

        if bound_presets:
            preset_list = "\n".join(f"  • {p}" for p in bound_presets)
            msg = (f"Delete voice:\n  {key}\n\n"
                   f"The following custom presets use this voice and will also be deleted:\n"
                   f"{preset_list}\n\n"
                   f"Delete the voice and all listed presets?")
            if not messagebox.askyesno("Confirm Delete", msg, parent=self.window):
                return
            for name in bound_presets:
                try:
                    os.remove(os.path.join(APP_PIPER_PRESETS_DIR, f"{name}.json"))
                except Exception:
                    pass
            self._refresh_piper_presets_list()
        else:
            if not messagebox.askyesno("Confirm Delete",
                                       f"Delete voice:\n{key}?",
                                       parent=self.window):
                return

        for ext in ('.onnx', '.onnx.json'):
            p = os.path.join(APP_PIPER_VOICES_DIR, f"{key}{ext}")
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
        if hasattr(self.app, '_piper_voice_cache'):
            self.app._piper_voice_cache.pop(key, None)
        self._set_status(f"Deleted: {key}", "green")
        self._refresh_lists()
        self._populate_voices_lists()
        if hasattr(self.app, "refresh_voice_dropdowns"):
            try:
                self.app.refresh_voice_dropdowns()
            except Exception:
                pass

    # --------------------------------------------------------------- helpers

    def _set_status(self, text, color="black"):
        self.status_label.config(text=text, fg=color)

    def _open_folder(self):
        try:
            os.startfile(APP_AI_VOICES_DIR)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}",
                                 parent=self.window)
