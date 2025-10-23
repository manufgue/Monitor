import tkinter as tk
from tkinter import ttk, messagebox
# Optional prettier UI with CustomTkinter. If not installed, fall back to standard ttk.
try:
    import customtkinter as ctk
    # prefer light mode
    ctk.set_appearance_mode('light')
    ctk.set_default_color_theme('blue')
    USE_CTK = True
    BaseWindow = ctk.CTk
except Exception:
    USE_CTK = False
    BaseWindow = tk.Tk
import requests
import threading
import json
import os
import time
import datetime
import re
import platform
try:
    import ctypes
    from ctypes import wintypes
except Exception:
    ctypes = None
from microfocus_logon_validate_logoff import microfocus_logon, microfocus_logoff
from validar_logon_logoff import logon_validate as local_logon_validate

# optional Pillow support for reliable GIF frame extraction
try:
    from PIL import Image, ImageTk
    _HAVE_PIL = True
except Exception:
    Image = None
    ImageTk = None
    _HAVE_PIL = False


class CompositeButton(tk.Frame):
    """A small composite widget that behaves like a button but is a Frame
    containing an optional image and a text label. It implements a minimal
    subset of the Button API used in this file: cget('text'), configure(...),
    pack(), bind(), winfo_height(), etc. This avoids issues where CTk/ttk
    buttons clip child widgets or don't accept place(in_=...)."""
    def __init__(self, parent, text, command, btn_fg='#4a4a4a', btn_text='#ffffff', *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self._command = command
        self._text = text
        self._state = 'normal'
        self._btn_fg = btn_fg
        self._btn_text = btn_text
        # inner frame to provide background and padding
        self._inner = tk.Frame(self, bg=self._btn_fg, bd=0, relief=tk.FLAT)
        self._inner.pack(fill=tk.BOTH, expand=True)
        self._img_label = None
        self._text_var = tk.StringVar(value=self._text)
        self._lbl = tk.Label(self._inner, textvariable=self._text_var, bg=self._btn_fg, fg=self._btn_text, padx=10, pady=6)
        self._lbl.pack(side='left')
        # bind clicks on both frame and label
        self._inner.bind('<Button-1>', self._on_click)
        self._lbl.bind('<Button-1>', self._on_click)

    def _on_click(self, event=None):
        if self._state == 'disabled':
            return
        try:
            if callable(self._command):
                self._command()
        except Exception:
            pass

    # Minimal Button-like API used by monitor.py
    def cget(self, key):
        if key == 'text':
            return self._text_var.get()
        return super().cget(key)

    def configure(self, **kwargs):
        if 'text' in kwargs:
            self._text_var.set(kwargs.get('text'))
            self._text = self._text_var.get()
        if 'state' in kwargs:
            st = kwargs.get('state')
            self._state = st
            if st == 'disabled':
                self._lbl.configure(fg='#888888')
            else:
                self._lbl.configure(fg=self._btn_text)
        # ignore other kwargs

    # expose pack/grid/place via inheritance; no override needed



class MonitorApp(BaseWindow):
    def __init__(self):
        super().__init__()
        self.title("ESMAC Active PCT Monitor")
        # try to set the window icon to Santander_logo.ico if available
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            base = os.getcwd()
        ico_path = os.path.join(base, 'Santander_logo.ico')
        try:
            if os.path.exists(ico_path):
                # On Windows, use wm_iconbitmap with .ico file
                try:
                    self.wm_iconbitmap(ico_path)
                except Exception:
                    # fallback: for CTk use iconphoto if available
                    try:
                        img = tk.PhotoImage(file=ico_path)
                        self.iconphoto(False, img)
                    except Exception:
                        pass
            else:
                # file not found; set a status message later when UI builds
                self._missing_icon = ico_path
        except Exception:
            # best-effort; don't block startup
            pass
        # start maximized but respect the taskbar (use Windows work area when possible)
        try:
            if platform.system() == 'Windows' and ctypes is not None:
                # SPI_GETWORKAREA -> fill usable desktop area (excludes taskbar)
                class RECT(ctypes.Structure):
                    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]
                SPI_GETWORKAREA = 0x0030
                rect = RECT()
                try:
                    ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
                    w = rect.right - rect.left
                    h = rect.bottom - rect.top
                    # set geometry to occupy the work area
                    self.geometry(f"{w}x{h}+{rect.left}+{rect.top}")
                except Exception:
                    # fallback to normal maximized state
                    try:
                        self.state('zoomed')
                    except Exception:
                        self.geometry("1200x800")
            else:
                # non-Windows: try normal maximized state and fallback
                try:
                    self.state('zoomed')
                except Exception:
                    self.geometry("1200x800")
        except Exception:
            try:
                self.state('zoomed')
            except Exception:
                self.geometry("1200x800")
        self.logon_headers = None
        self.session_cookie = None
        # store session cookies per host so we can reuse them when querying many hosts
        # mapping: host -> cookie string
        self.session_cookies = {}
        # presets and host->region mappings
        # default_presets: static list shipped with the GUI. Keep a copy so we can
        # merge saved mappings with these defaults (so hosts like 10.150.112.157
        # are always available in the combobox even if not saved in mapping).
        self.default_presets = [
            "10.150.112.142",
            "10.150.112.143",
            "10.150.112.147",
            "10.150.112.148",
            "10.150.112.149",
            "10.150.112.150",
            "10.150.112.151",
            "10.150.112.152",
            "10.150.112.153",
            "10.150.112.154",
            "10.150.112.155",
            "10.150.112.156",
            "10.150.112.166",
            # include the common PROD example host so it's visible by default
            "10.150.112.157",
        ]
        # start host_presets from the defaults; we'll merge with saved mapping keys below
        self.host_presets = list(self.default_presets)
        # Provide a small default mapping for common hosts so the GUI shows regions
        # immediately even if host_mapping.json is missing or empty.
        self.default_host_mapping = {
            "10.150.112.157": {
                "regions": ["GRAVP168", "GRAVP169", "GRAVP170"],
                "port": 10086,
            }
        }
        self.host_mapping = {}  # ip -> {"region": "GRAVPxxx", "port": 10086}
        self._load_host_mapping()
        # Merge saved mapping keys with the default presets so both appear in the combobox.
        # Also merge in any default_host_mapping entries if mapping file is empty.
        try:
            # ensure host_mapping has the default entries for known hosts
            for k, v in self.default_host_mapping.items():
                if k not in self.host_mapping:
                    self.host_mapping[k] = v
            combined = sorted(set(self.default_presets) | set(self.host_mapping.keys()))
            # remove any default IPs in the 10.150.112.* range from the dropdown
            self.host_presets = [h for h in combined if not h.startswith('10.150.112.')]
        except Exception:
            # fallback: use defaults
            self.host_presets = list(self.default_presets)
        self._build_ui()
        

    def _build_ui(self):
        # Widget factories: prefer CustomTkinter when available, otherwise use ttk
        Frame = ctk.CTkFrame if USE_CTK else ttk.Frame
        Label = ctk.CTkLabel if USE_CTK else ttk.Label
        EntryW = ctk.CTkEntry if USE_CTK else ttk.Entry
        ButtonW = ctk.CTkButton if USE_CTK else ttk.Button
        CheckW = ctk.CTkCheckBox if USE_CTK else ttk.Checkbutton
        ComboBox = ctk.CTkComboBox if USE_CTK and hasattr(ctk, 'CTkComboBox') else ttk.Combobox

        # sensible width values: CustomTkinter measures width in pixels; ttk uses char-count
        combo_width = 220 if USE_CTK else 18
        entry_width = 200 if USE_CTK else 18

        # button colors
        btn_fg = '#4a4a4a'
        btn_hover = '#606060'
        btn_text = '#ffffff'

        # helper to create a styled button for both CTk and ttk
        def mk_button(parent, text, command, **kwargs):
            if USE_CTK:
                try:
                    return ctk.CTkButton(parent, text=text, command=command, fg_color=btn_fg, hover_color=btn_hover, text_color=btn_text, **kwargs)
                except Exception:
                    return ctk.CTkButton(parent, text=text, command=command, **kwargs)
            else:
                try:
                    return ttk.Button(parent, text=text, command=command, style='Dark.TButton', **kwargs)
                except Exception:
                    return ttk.Button(parent, text=text, command=command, **kwargs)

        # Create main frames used by the UI: 'frm' is the primary content frame and 'inputs' holds the top input row.
        try:
            frm = Frame(self)
            frm.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        except Exception:
            frm = Frame(self)
            frm.pack(fill=tk.BOTH, expand=True)

        # Authentication controls: Usuario / Password and Logon/Logoff buttons
        try:
            auth = Frame(frm)
            auth.pack(fill=tk.X, padx=6, pady=(0,6))
        except Exception:
            auth = Frame(frm)
            auth.pack(fill=tk.X)

        # StringVars for user and password used across the app
        try:
            self.user_var = tk.StringVar(value='')
        except Exception:
            self.user_var = tk.StringVar()
        try:
            self.pass_var = tk.StringVar(value='')
        except Exception:
            self.pass_var = tk.StringVar()

        # Usuario label + entry
        try:
            Label(auth, text='Usuario:').grid(row=0, column=0, sticky=tk.W, padx=2, pady=2)
            if USE_CTK and EntryW is ctk.CTkEntry:
                self.user_entry = EntryW(auth, textvariable=self.user_var, width=160)
                self.user_entry.grid(row=0, column=1, sticky='w', padx=6, pady=2)
            else:
                self.user_entry = EntryW(auth, textvariable=self.user_var, width=18)
                self.user_entry.grid(row=0, column=1, sticky=tk.W, padx=6, pady=2)
        except Exception:
            pass

        # Password label + entry
        try:
            Label(auth, text='Password:').grid(row=0, column=2, sticky=tk.W, padx=6, pady=2)
            if USE_CTK and EntryW is ctk.CTkEntry:
                self.pass_entry = EntryW(auth, textvariable=self.pass_var, width=160, show='*')
                self.pass_entry.grid(row=0, column=3, sticky='w', padx=6, pady=2)
            else:
                self.pass_entry = EntryW(auth, textvariable=self.pass_var, width=18, show='*')
                self.pass_entry.grid(row=0, column=3, sticky=tk.W, padx=6, pady=2)
        except Exception:
            pass

        # Logon / Logoff buttons
        try:
            self.btn_logon = mk_button(auth, text='Logon', command=lambda: self.login())
            self.btn_logon.grid(row=0, column=4, padx=6)
        except Exception:
            try:
                self.btn_logon = mk_button(auth, text='Logon', command=lambda: self.login())
                self.btn_logon.pack(side='left', padx=6)
            except Exception:
                pass
        try:
            self.btn_logoff = mk_button(auth, text='Logoff', command=self.logoff)
            self.btn_logoff.grid(row=0, column=5, padx=6)
        except Exception:
            try:
                self.btn_logoff = mk_button(auth, text='Logoff', command=self.logoff)
                self.btn_logoff.pack(side='left', padx=6)
            except Exception:
                pass

        try:
            inputs = Frame(frm)
            inputs.pack(fill=tk.X, padx=6, pady=6)
        except Exception:
            inputs = Frame(frm)
            inputs.pack(fill=tk.X)

        Label(inputs, text="Host:").grid(row=0, column=0, sticky=tk.W, padx=2, pady=2)
        # start with empty Host field; user will select or type one
        self.host_var = tk.StringVar(value="")
        # Host combobox with presets (user can also type)
        if USE_CTK and hasattr(ctk, 'CTkComboBox'):
            self.host_combo = ctk.CTkComboBox(inputs, values=self.host_presets, width=combo_width)
            # CTkComboBox uses a command callback for selection
            self.host_combo.grid(row=0, column=1, sticky='w', padx=6, pady=4)
            # when selection changes, update host_var and run selection logic
            self.host_combo.configure(command=lambda val: (self.host_var.set(val), self._on_host_selected()))
        else:
            self.host_combo = ttk.Combobox(inputs, textvariable=self.host_var, values=self.host_presets, width=combo_width)
            self.host_combo.grid(row=0, column=1, sticky=tk.W, padx=6, pady=4)
            self.host_combo.bind('<<ComboboxSelected>>', lambda e: self._on_host_selected())
            self.host_combo.bind('<FocusOut>', lambda e: self._on_host_entry_changed())

        Label(inputs, text="Port:").grid(row=0, column=2, sticky=tk.W, padx=6)
        self.port_var = tk.StringVar(value="10086")
        # Port entry
        if USE_CTK and EntryW is ctk.CTkEntry:
            self.port_entry = EntryW(inputs, textvariable=self.port_var, width=80)
            self.port_entry.grid(row=0, column=3, sticky='w', padx=6, pady=4)
        else:
            self.port_entry = EntryW(inputs, textvariable=self.port_var, width=8)
            self.port_entry.grid(row=0, column=3, sticky=tk.W, padx=6, pady=4)

        # hide host and port controls until login performed
        try:
            self.host_combo.grid_remove()
            self.port_entry.grid_remove()
            # also hide their labels (they are at fixed grid positions)
            try:
                for child in inputs.grid_slaves(row=0, column=0):
                    child.grid_remove()
                for child in inputs.grid_slaves(row=0, column=2):
                    child.grid_remove()
            except Exception:
                pass
        except Exception:
            pass

        # Canal and Site (read-only) shown for selected host
        Label(inputs, text="Canal:").grid(row=0, column=4, sticky=tk.W, padx=6)
        self.canal_var = tk.StringVar(value="")
        # Canal (read-only)
        if USE_CTK and EntryW is ctk.CTkEntry:
            self.canal_entry = EntryW(inputs, textvariable=self.canal_var, width=entry_width)
            self.canal_entry.grid(row=0, column=5, sticky='w', padx=6, pady=4)
            try:
                self.canal_entry.configure(state='readonly')
            except Exception:
                pass
        else:
            self.canal_entry = EntryW(inputs, textvariable=self.canal_var, width=12, state='readonly')
            self.canal_entry.grid(row=0, column=5, sticky=tk.W, padx=6, pady=4)

        Label(inputs, text="Site:").grid(row=0, column=6, sticky=tk.W, padx=6)
        self.site_var = tk.StringVar(value="")
        # Site (read-only)
        if USE_CTK and EntryW is ctk.CTkEntry:
            self.site_entry = EntryW(inputs, textvariable=self.site_var, width=160)
            self.site_entry.grid(row=0, column=7, sticky='w', padx=6, pady=4)
            try:
                self.site_entry.configure(state='readonly')
            except Exception:
                pass
        else:
            self.site_entry = EntryW(inputs, textvariable=self.site_var, width=10, state='readonly')
            self.site_entry.grid(row=0, column=7, sticky=tk.W, padx=6, pady=4)

        Label(inputs, text="Region:").grid(row=0, column=8, sticky=tk.W, padx=6, pady=4)
        # region_var holds the selected region (single string) - start empty
        self.region_var = tk.StringVar(value="")
        # Add a combobox to display multiple associated regions for the selected host
        if USE_CTK and hasattr(ctk, 'CTkComboBox'):
            self.regions_combo = ctk.CTkComboBox(inputs, values=[], width=combo_width)
            self.regions_combo.grid(row=0, column=9, sticky='w', padx=6, pady=4)
            self.regions_combo.configure(command=lambda val: self.region_var.set(val))
            # CTkComboBox sometimes shows a default label like 'CTkComboBox' — clear it
            try:
                self.regions_combo.set('')
            except Exception:
                # If set is not supported, ensure the bound variable is empty
                try:
                    self.region_var.set('')
                except Exception:
                    pass
        else:
            self.regions_combo = ttk.Combobox(inputs, textvariable=self.region_var, values=[], width=combo_width)
            self.regions_combo.grid(row=0, column=9, sticky=tk.W, padx=6, pady=4)
            self.regions_combo.bind('<<ComboboxSelected>>', lambda e: None)

    # mapping editor button removed per request

        # Options (removed: statusCodes / resFilter / groupFilter)

        # Buttons (use the same Frame factory so background matches the rest of UI)
        btns = Frame(frm)
        try:
            btns.pack(fill=tk.X, pady=6)
        except Exception:
            # CTk Frame may use pack similarly; fallback to pack anyway
            btns.pack(fill=tk.X, pady=6)
        # Action buttons (styled)
        self.btn_refresh = mk_button(btns, text="Refrescar (Enter)", command=self.refresh)
        self.btn_refresh.pack(side='left', padx=6)
        self.btn_clear = mk_button(btns, text="Limpiar tabla", command=self.clear_table)
        self.btn_clear.pack(side='left', padx=6)
        # create a small container so we can reliably show an image/spinner left of the button
        try:
            self.btn_consultar_frame = Frame(btns)
            self.btn_consultar_frame.pack(side='left', padx=6)
        except Exception:
            self.btn_consultar_frame = btns
        # Create an inner container for the Consultar total button and its spinners/gif
        try:
            self.btn_consultar_inner = Frame(self.btn_consultar_frame)
            self.btn_consultar_inner.pack(side='left')
        except Exception:
            self.btn_consultar_inner = self.btn_consultar_frame

        # Add an indeterminate spinner/progressbar inside the inner container (hidden by default)
        try:
            self.spinner = ttk.Progressbar(self.btn_consultar_inner, mode='indeterminate', length=80)
            try:
                self.spinner.pack_forget()
            except Exception:
                pass
        except Exception:
            self.spinner = None

        # small circular spinner implemented on a Canvas (works without external assets)
        try:
            cs = 16
            self.circle_spinner = tk.Canvas(self.btn_consultar_inner, width=cs, height=cs, highlightthickness=0)
            try:
                arc_color = btn_text if 'btn_text' in locals() else '#555555'
                self._circle_arc = self.circle_spinner.create_arc(2, 2, cs-2, cs-2, start=0, extent=90, style='arc', width=3, outline=arc_color)
            except Exception:
                self._circle_arc = None
        except Exception:
            self.circle_spinner = None

        # Attempt to load an animated GIF spinner from GUIS/loading_spinner_small.gif or loading_spinner.gif
        try:
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            except Exception:
                base_dir = os.getcwd()
            gif_small = os.path.join(base_dir, 'GUIS', 'loading_spinner_small.gif')
            gif_default = os.path.join(base_dir, 'GUIS', 'loading_spinner.gif')
            gif_path = gif_small if os.path.exists(gif_small) else (gif_default if os.path.exists(gif_default) else None)
            if gif_path:
                try:
                    self._gif_frames = []
                    if _HAVE_PIL:
                        try:
                            pil_gif = Image.open(gif_path)
                            for frame in range(getattr(pil_gif, 'n_frames', 1)):
                                try:
                                    pil_gif.seek(frame)
                                    frame_image = ImageTk.PhotoImage(pil_gif.convert('RGBA'))
                                    self._gif_frames.append(frame_image)
                                except Exception:
                                    break
                        except Exception:
                            self._gif_frames = []
                    if not self._gif_frames:
                        try:
                            idx = 0
                            while True:
                                try:
                                    f = tk.PhotoImage(file=gif_path, format=f'gif -index {idx}')
                                    self._gif_frames.append(f)
                                    idx += 1
                                except Exception:
                                    break
                        except Exception:
                            self._gif_frames = []
                    if self._gif_frames:
                        try:
                            self.gif_label = tk.Label(self.btn_consultar_inner, image=self._gif_frames[0], bd=0)
                        except Exception:
                            self.gif_label = tk.Label(self, image=self._gif_frames[0], bd=0)
                        self._gif_idx = 0
                        self._gif_running = False
                        try:
                            self._gif_path = gif_path
                        except Exception:
                            self._gif_path = None
                    else:
                        self.gif_label = None
                        self._gif_frames = None
                except Exception:
                    self.gif_label = None
                    self._gif_frames = None
            else:
                self.gif_label = None
                self._gif_frames = None
        except Exception:
            self.gif_label = None
            self._gif_frames = None

        # create the Consultar total button inside the inner container so spinner/gif can sit next to it
        self.btn_consultar_todos = mk_button(self.btn_consultar_inner, text="Consultar total", command=self.consultar_todos)
        self.btn_consultar_todos.pack(side='left')

        # Treeview
        cols = ("PCTName", "group", "PCTSec", "PCTCnt")
        self.tree = ttk.Treeview(frm, columns=cols, show="headings")
        for c in cols:
            # display 'group' column header as 'Group' (capitalized) while keeping internal id 'group'
            display = 'Group' if c == 'group' else c
            self.tree.heading(c, text=display)
            self.tree.column(c, width=260 if c == "PCTName" else 120, anchor=tk.W)

        # Apply light Treeview styling (works for ttk; CTk background remains as set)
        try:
            s = ttk.Style()
            # ensure a theme that allows background changes
            s.theme_use('clam')
            s.configure('Treeview', background='#ffffff', foreground='#000000', fieldbackground='#ffffff', rowheight=24)
            s.configure('Treeview.Heading', background='#f0f0f0', foreground='#000000')
            s.map('Treeview', background=[('selected', '#cfe2ff')], foreground=[('selected', '#000000')])
        except Exception:
            pass

        self.tree.pack(fill=tk.BOTH, expand=True)
        # configure zebra stripe tags (light)
        try:
            self.tree.tag_configure('even', background='#ffffff')
            self.tree.tag_configure('odd', background='#f7f7f7')
        except Exception:
            pass

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        if USE_CTK:
            # CTk doesn't have a direct status label; use ttk Label for compatibility
            status = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        else:
            status = ttk.Label(self, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status.pack(fill=tk.X, side=tk.BOTTOM)

        # Bind Enter to refresh
        self.bind('<Return>', lambda e: self.refresh())
        # Start with UI controls disabled until user logs in
        try:
            self._set_ui_logged_in(False)
        except Exception:
            pass

    

    def _apply_preset(self):
        # legacy/unused - kept for compatibility but no preset widget exists
        return

    def _host_mapping_file_path(self):
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except Exception:
            base = os.getcwd()
        return os.path.join(base, 'host_mapping.json')

    def _on_host_selected(self):
        """When the user selects a host from the combobox, try to fill region and port
        from the loaded host_mapping if present."""
        ip = self.host_var.get().strip()
        # normalize compact IPs if the user entered digits without dots
        norm = self._normalize_compact_ip(ip)
        if norm and norm != ip:
            self.host_var.set(norm)
            ip = norm
        if not ip:
            return
        mapping = self.host_mapping.get(ip)
        if not mapping:
            return
        # mapping may contain 'regions' (list) or legacy 'region' (single)
        regions = mapping.get('regions') or ([mapping.get('region')] if mapping.get('region') else [])
        port = mapping.get('port')
        canal = mapping.get('canal') or mapping.get('Canal') or ''
        site = mapping.get('site') or mapping.get('Site') or ''
        try:
            # include an 'All' option that represents aggregating across all regions for this host
            display_regions = ['All'] + regions if regions else []
            # update regions_combo values and ensure the selected region is synchronized
            try:
                if USE_CTK and hasattr(ctk, 'CTkComboBox') and isinstance(self.regions_combo, ctk.CTkComboBox):
                    # CTkComboBox uses configure(values=...) and set(...) methods
                    try:
                        self.regions_combo.configure(values=display_regions)
                    except Exception:
                        # older CTk versions may store values differently
                        try:
                            self.regions_combo._set_values(display_regions)
                        except Exception:
                            pass
                else:
                    # ttk.Combobox
                    try:
                        self.regions_combo['values'] = display_regions
                    except Exception:
                        pass

                # If the mapping provides regions, select the first one explicitly and
                # synchronize both the widget and the bound StringVar. This avoids
                # a stale region remaining selected (which can produce 404 on refresh).
                if regions:
                    sel = regions[0]
                    # try several approaches to set the combobox value (CTk/ttk differences)
                    try:
                        if hasattr(self.regions_combo, 'set'):
                            try:
                                self.regions_combo.set(sel)
                            except Exception:
                                pass
                        # also set the bound variable
                        try:
                            self.region_var.set(sel)
                        except Exception:
                            pass
                        # if ttk, also generate the virtual event so any bindings update
                        try:
                            if not (USE_CTK and hasattr(ctk, 'CTkComboBox')):
                                try:
                                    self.regions_combo.event_generate('<<ComboboxSelected>>')
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    except Exception:
                        # best-effort fallback
                        try:
                            self.region_var.set(sel)
                        except Exception:
                            pass
                else:
                    # no regions for this host: clear selection to avoid using a stale region
                    try:
                        self.region_var.set('')
                    except Exception:
                        pass
                    try:
                        if hasattr(self.regions_combo, 'set'):
                            try:
                                self.regions_combo.set('')
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass
        except Exception:
            pass
        # populate canal/site fields
        try:
            if canal:
                self.canal_var.set(canal)
            else:
                self.canal_var.set('')
        except Exception:
            pass
        try:
            if site:
                self.site_var.set(site)
            else:
                self.site_var.set('')
        except Exception:
            pass
        if port:
            try:
                self.port_var.set(str(port))
            except Exception:
                pass

    def _normalize_compact_ip(self, text: str) -> str:
        """If text is a compact digits-only IP like 10150112157, convert to dotted
        form 10.150.112.157. If not convertible, return empty string.
        Simple heuristic: if length is 11-12 and only digits, split into 4 parts
        taking typical octet sizes. This is a best-effort helper; if it can't
        produce a valid dotted IP, returns empty string.
        """
        if not text:
            return ''
        s = text.strip()
        if not s.isdigit():
            return ''
        # try to split into 4 octets by assuming each octet can be 1-3 digits
        # we'll attempt a few heuristics; prefer splitting into 4 groups of similar size
        L = len(s)
        if L < 4 or L > 12:
            return ''
        # distribute lengths as evenly as possible into 4 parts
        base = L // 4
        rem = L % 4
        sizes = [base + (1 if i < rem else 0) for i in range(4)]
        idx = 0
        parts = []
        for sz in sizes:
            part = s[idx:idx+sz]
            if not part:
                return ''
            val = int(part)
            if val < 0 or val > 255:
                return ''
            parts.append(str(val))
            idx += sz
        dotted = '.'.join(parts)
        return dotted

    def _parse_count(self, cnt_raw) -> int:
        """Robustly parse a count that may include thousand separators like '4,818' or spaces.
        Returns an int (0 on failure).
        """
        if cnt_raw is None:
            return 0
        s = str(cnt_raw).strip()
        if not s:
            return 0
        # remove common thousand separators and spaces
        s2 = re.sub(r'[\,\s]', '', s)
        try:
            return int(s2)
        except Exception:
            try:
                return int(float(s2))
            except Exception:
                # fallback: keep only digits
                digits = re.sub(r'\D', '', s)
                try:
                    return int(digits) if digits else 0
                except Exception:
                    return 0

    def _on_host_entry_changed(self):
        # User finished editing host entry — normalize if necessary and trigger selection logic
        ip = self.host_var.get().strip()
        norm = self._normalize_compact_ip(ip)
        if norm and norm != ip:
            self.host_var.set(norm)
            # ensure combobox values include this ip
            try:
                values = list(self.host_combo['values'])
                if norm not in values:
                    values.append(norm)
                    self.host_combo['values'] = sorted(values)
            except Exception:
                pass
        # call host selected logic to populate regions/port if mapping exists
        try:
            self._on_host_selected()
        except Exception:
            pass

    def _load_host_mapping(self):
        path = self._host_mapping_file_path()
        if not os.path.exists(path):
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if isinstance(data, dict):
                self.host_mapping.update(data)
        except Exception:
            pass

    def _save_host_mapping(self):
        path = self._host_mapping_file_path()
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.host_mapping, f, indent=2)
            self.status_var.set(f"Host mapping guardado ({len(self.host_mapping)} entradas)")
        except Exception as e:
            self.status_var.set(f"Error guardando mapping: {e}")

    def _open_mapping_editor(self):
        # simple modal to paste CSV lines: ip,region,port
        top = tk.Toplevel(self)
        top.title('Editar host mapping')
        top.geometry('600x300')
        lbl = ttk.Label(top, text='Pegue líneas CSV: ip,region,port (uno por línea). Luego presione Guardar')
        lbl.pack(fill=tk.X, padx=8, pady=6)
        txt = tk.Text(top)
        # prefill with current mapping lines
        lines = []
        for ip, m in self.host_mapping.items():
            # support multiple regions per host, stored as a list under 'regions'
            regions = m.get('regions') or ([m.get('region')] if m.get('region') else [])
            regions_field = '|'.join(regions) if regions else ''
            lines.append(f"{ip},{regions_field},{m.get('port',10086)}")
        txt.insert('1.0', '\n'.join(lines))
        txt.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        def _save_from_text():
            content = txt.get('1.0', tk.END).strip()
            newmap = {}
            for line in content.splitlines():
                # Try to accept multiple formats, including Excel-like rows.
                if not line.strip():
                    continue
                # detect separator: prefer tab, then semicolon, then comma
                if '\t' in line:
                    tokens = [t.strip() for t in line.split('\t') if t.strip()]
                elif ';' in line and ',' not in line:
                    tokens = [t.strip() for t in line.split(';') if t.strip()]
                else:
                    tokens = [t.strip() for t in line.split(',') if t.strip()]

                # heuristics to find ip/host, region(s), and port
                ip_candidate = None
                regions_field = ''
                port = None

                # search tokens for IP-like or hostname-like token
                ip_re = re.compile(r"^\d{1,3}(?:\.\d{1,3}){3}$")
                host_re = re.compile(r"^[A-Za-z0-9\-_.]+$")
                region_re = re.compile(r"^[A-Z]{2,6}[A-Z0-9]*")
                for t in tokens:
                    if ip_re.match(t):
                        ip_candidate = t
                        continue
                    # hostnames like iacvm1777.ar.bsch
                    if (not ip_candidate) and ('iacvm' in t or '.' in t and any(c.isalpha() for c in t)):
                        ip_candidate = t
                        continue
                # try to find port (4-5 digit number)
                for t in reversed(tokens):
                    if t.isdigit() and 1 < len(t) <= 5:
                        port = int(t)
                        break

                # try to find region token (CICS column) - look for tokens starting with GRAV or similar
                for t in tokens:
                    if region_re.match(t) and ('GRAV' in t or 'GR' in t or t.isupper()):
                        regions_field = t
                        break

                # fallback: if tokens length >=3, assume 2nd token is region-like
                if not regions_field and len(tokens) >= 2:
                    regions_field = tokens[1]

                # normalize ip candidate if compact digits
                if ip_candidate:
                    try:
                        ip_candidate = self._normalize_compact_ip(ip_candidate) or ip_candidate
                    except Exception:
                        pass

                # split regions_field by | or ; or / or space if multiple
                regions = []
                if regions_field:
                    regions = [r.strip() for r in re.split(r"[|;/\\]", regions_field) if r.strip()]
                if not regions and regions_field:
                    regions = [regions_field.strip()]

                if not port:
                    port = 10086

                if not ip_candidate:
                    # cannot determine host key; skip line
                    continue
                newmap[ip_candidate] = {'regions': regions, 'port': port}
            self.host_mapping = newmap
            self._save_host_mapping()
            # update presets combobox values (merge defaults with saved mapping keys)
            try:
                combined = sorted(set(self.default_presets) | set(self.host_mapping.keys()))
                # filter out the 10.150.112.* defaults from the combobox values
                self.host_presets = [h for h in combined if not h.startswith('10.150.112.')]
            except Exception:
                self.host_presets = list(self.default_presets)
            try:
                # update the combobox values so new hosts appear in the Host: field
                self.host_combo['values'] = self.host_presets
                # if there is at least one host, select the first one
                if self.host_presets:
                    self.host_combo.set(self.host_presets[0])
                    # also populate region/port for the selected host
                    self._on_host_selected()
            except Exception:
                pass
            top.destroy()

        btn = ttk.Button(top, text='Guardar', command=_save_from_text)
        btn.pack(pady=6)

        # small debug button to view current mapping quickly
        dbg = ttk.Button(top, text='Mostrar mapping', command=lambda: messagebox.showinfo('Mapping', json.dumps(self.host_mapping, indent=2)))
        dbg.pack(pady=2)

    def consultar_todos(self):
        # Query all mapped hosts and aggregate PCTCnt across ALL regions of ALL hosts
        # show busy indicator on the Consultar todos button
        if not self.host_mapping:
            messagebox.showinfo('Sin mapping', 'No hay mapping de hosts. Use el editor de mapping o agregue host_mapping.json')
            return

        # set busy indicator and run heavy work in background thread so the UI remains responsive
        self._set_busy(True)

        def _start():
            try:
                result = self._consultar_todos_worker()
            except Exception as e:
                result = {'error': str(e)}
            # schedule completion handler in main thread
            try:
                self.after(0, lambda: self._consultar_todos_done(result))
            except Exception:
                pass

        t = threading.Thread(target=_start, daemon=True)
        t.start()


    def _set_busy(self, busy: bool):
        """Toggle busy state for the Consultar todos button. When busy, disable the button
        and start a small animation in its text. When not busy, restore the original text.
        """
        try:
            if not hasattr(self, 'btn_consultar_todos') or self.btn_consultar_todos is None:
                return
            if busy:
                try:
                    self._old_btn_text = self.btn_consultar_todos.cget('text')
                except Exception:
                    self._old_btn_text = 'Consultar todos'
                try:
                    self.btn_consultar_todos.configure(state='disabled')
                except Exception:
                    try:
                        self.btn_consultar_todos.configure(text='Consultando...')
                    except Exception:
                        pass
                # show and start spinner if available
                try:
                    if hasattr(self, 'spinner') and self.spinner:
                        try:
                            # pack the spinner immediately to the left of the Consultar todos button
                            try:
                                self.spinner.pack(side='left', padx=6, before=self.btn_consultar_todos)
                            except Exception:
                                # older tkinter versions may not support 'before' or btn may be CTk; fallback
                                self.spinner.pack(side='left', padx=6)
                        except Exception:
                            pass
                        try:
                            self.spinner.start(10)
                        except Exception:
                            pass
                except Exception:
                    pass
                # start canvas circular spinner
                # If we have an animated GIF, prefer it (place it inside the button)
                try:
                    if getattr(self, 'gif_label', None) and getattr(self, '_gif_frames', None):
                        try:
                            # try to pack the gif label into the same container as the buttons, immediately before the Consultar button
                            placed = False
                            try:
                                placed = self._place_gif_label()
                            except Exception:
                                placed = False
                            # if not placed, schedule a retry after layout
                            if not placed:
                                try:
                                    self.after(50, self._place_gif_label)
                                except Exception:
                                    pass
                            # brief debug info in status bar to help diagnose visibility
                            try:
                                p = getattr(self, '_gif_path', None)
                                nfr = len(self._gif_frames) if getattr(self, '_gif_frames', None) else 0
                                self.status_var.set(f"Spinner GIF started ({p}) frames={nfr}")
                            except Exception:
                                pass
                            self._gif_running = True
                            self._gif_idx = 0
                            self._animate_gif()
                        except Exception:
                            # if pack fails, fallback to canvas spinner
                            self._start_circle_spinner()
                    else:
                        self._start_circle_spinner()
                except Exception:
                    pass
                # no external busy label; spinner/circle handle visual feedback
                # start animation ticker
                self._busy_dots = 0
                self._busy_running = True
                self.after(200, self._busy_tick)
            else:
                self._busy_running = False
                try:
                    self.btn_consultar_todos.configure(state='normal')
                except Exception:
                    pass
                try:
                    if hasattr(self, '_old_btn_text'):
                        self.btn_consultar_todos.configure(text=self._old_btn_text)
                except Exception:
                    pass
                # no external busy label to hide
            # stop and hide spinner if available
            try:
                if hasattr(self, 'spinner') and self.spinner:
                    try:
                        self.spinner.stop()
                    except Exception:
                        pass
                    try:
                        self.spinner.pack_forget()
                    except Exception:
                        pass
            except Exception:
                pass
            # stop canvas circular spinner
            # stop gif animation and hide label if used
            try:
                if getattr(self, '_gif_running', False) and getattr(self, 'gif_label', None):
                    try:
                        self._gif_running = False
                        # try to remove via pack_forget first
                        try:
                            self.gif_label.pack_forget()
                        except Exception:
                            try:
                                self.gif_label.place_forget()
                            except Exception:
                                pass
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                self._stop_circle_spinner()
            except Exception:
                pass
        except Exception:
            pass

    def _busy_tick(self):
        if not getattr(self, '_busy_running', False):
            return
        try:
            self._busy_dots = (getattr(self, '_busy_dots', 0) + 1) % 4
            dots = '.' * self._busy_dots
            try:
                self.btn_consultar_todos.configure(text=f"Consultando{dots}")
            except Exception:
                pass
            # no external busy label to update; circular spinner provides feedback
            self.after(400, self._busy_tick)
        except Exception:
            pass

    def _start_circle_spinner(self):
        try:
            if not hasattr(self, 'circle_spinner') or not self.circle_spinner:
                return
            try:
                # attempt to place the small canvas inside the Consultar todos button so it appears to the left of the text
                try:
                    # compute a small offset (left side inside button)
                    bx = 4
                    by = int((self.btn_consultar_todos.winfo_height() - self.circle_spinner.winfo_reqheight()) / 2)
                    # use place(in_=...) to position relative to the button widget
                    self.circle_spinner.place(in_=self.btn_consultar_todos, x=bx, y=max(0, by))
                except Exception:
                    # fallback: pack to the left of the button's parent
                    try:
                        self.circle_spinner.pack(side='left', padx=4, before=self.btn_consultar_todos)
                    except Exception:
                        self.circle_spinner.pack(side='left', padx=4)
            except Exception:
                pass
            self._circle_angle = 0
            self._circle_running = True
            # ensure the animation loop starts
            self._circle_step()
        except Exception:
            pass

    def _stop_circle_spinner(self):
        try:
            self._circle_running = False
        except Exception:
            pass
        try:
            if hasattr(self, 'circle_spinner') and self.circle_spinner:
                try:
                    # hide whichever geometry manager was used
                    self.circle_spinner.place_forget()
                except Exception:
                    try:
                        self.circle_spinner.pack_forget()
                    except Exception:
                        pass
        except Exception:
            pass

    def _circle_step(self):
        try:
            if not getattr(self, '_circle_running', False):
                return
            self._circle_angle = (getattr(self, '_circle_angle', 0) + 15) % 360
            try:
                if getattr(self, '_circle_arc', None) is not None:
                    self.circle_spinner.itemconfigure(self._circle_arc, start=self._circle_angle)
            except Exception:
                pass
            self.after(50, self._circle_step)
        except Exception:
            pass

    def _animate_gif(self):
        try:
            if not getattr(self, '_gif_running', False) or not getattr(self, '_gif_frames', None):
                return
            try:
                self._gif_idx = (getattr(self, '_gif_idx', 0) + 1) % len(self._gif_frames)
                frame = self._gif_frames[self._gif_idx]
                try:
                    self.gif_label.configure(image=frame)
                except Exception:
                    pass
            except Exception:
                pass
            # schedule next frame
            self.after(80, self._animate_gif)
        except Exception:
            pass

    def _place_gif_label(self) -> bool:
        """Attempt to pack the gif_label before the Consultar button. Returns True on success."""
        try:
            if not getattr(self, 'gif_label', None) or not getattr(self, 'btn_consultar_todos', None):
                return False
            try:
                self.gif_label.pack(side='left', padx=4, before=self.btn_consultar_todos)
                try:
                    self.gif_label.lift()
                except Exception:
                    pass
                return True
            except Exception:
                try:
                    self.gif_label.pack(side='left', padx=4)
                    try:
                        self.gif_label.lift()
                    except Exception:
                        pass
                    return True
                except Exception:
                    # fallback: place the gif_label over the button using absolute coords
                    try:
                        # get button position relative to root window
                        bx = self.btn_consultar_todos.winfo_rootx() - self.winfo_rootx()
                        by = self.btn_consultar_todos.winfo_rooty() - self.winfo_rooty()
                        bw = self.btn_consultar_todos.winfo_width()
                        bh = self.btn_consultar_todos.winfo_height()
                        # place image 4px from left inside button vertically centered
                        x = bx + 4
                        y = by + max(0, (bh - self.gif_label.winfo_reqheight()) // 2)
                        self.gif_label.place(x=x, y=y)
                        try:
                            self.gif_label.lift()
                        except Exception:
                            pass
                        return True
                    except Exception:
                        return False
        except Exception:
            return False

    

    def _get_session_cookie_for_host(self, host: str, port: str, interactive: bool = True) -> str:
        """Attempt to obtain a session cookie for a specific host:port.
        If interactive is True, use messageboxes for errors; otherwise be silent.
        Returns cookie string on success or None on failure.
        """
        base_url = f"http://{host}:{port}"
        user = self.user_var.get().strip()
        pw = self.pass_var.get().strip()
        if not user or not pw:
            if interactive:
                messagebox.showwarning("Missing", "Usuario y password son requeridos para iniciar sesión")
            return None
        try:
            headers, cookie = microfocus_logon(base_url, user, pw)
        except SystemExit:
            if interactive:
                messagebox.showerror("Logon failed", "Logon fallido")
            return None
        except Exception as e:
            if interactive:
                messagebox.showerror("Logon error", str(e))
            return None
        # store cookie for this host for reuse
        try:
            if cookie:
                self.session_cookies[host] = cookie
        except Exception:
            pass
        # also update last interactive session fields
        if interactive:
            self.logon_headers = headers
            self.session_cookie = cookie
            self.status_var.set("Logon exitoso")
        return cookie

    def _consultar_todos_worker(self):
        """Background worker that queries all hosts/regions and aggregates counts.
        Returns a dict with keys: total_sum, total_calls, agg_pcts, agg_by_region, agg_by_host, error(optional)
        This function MUST NOT call GUI methods (messagebox, status_var, tree insertion)."""
        total_sum = 0
        total_calls = 0
        agg_pcts = {}
        agg_by_region = {}
        agg_by_host = {}

        for host_key, m in list(self.host_mapping.items()):
            port = m.get('port', 10086)
            regions = m.get('regions') or ([m.get('region')] if m.get('region') else [])
            if not regions:
                continue

            # get or try to obtain a cookie non-interactively
            cookie_val = None
            try:
                cookie_val = self.session_cookies.get(host_key)
            except Exception:
                cookie_val = None
            if not cookie_val:
                try:
                    if self.user_var.get().strip() and self.pass_var.get().strip():
                        cookie_val = self._get_session_cookie_for_host(host_key, port, interactive=False)
                except Exception:
                    cookie_val = None

            for region in regions:
                url = f"http://{host_key}:{port}/native/v1/regions/{host_key}/86/{region}/active/pct"
                try:
                    headers = {"Content-Type": "application/json", "Origin": f"http://{host_key}:{port}", "X-Requested-With": "XMLHttpRequest"}
                    cookies = {"ESAdmin-Cookie": cookie_val} if cookie_val else None
                    resp = requests.get(url, timeout=12, headers=headers, cookies=cookies)
                except Exception:
                    continue

                total_calls += 1
                if resp.status_code == 401:
                    # try a non-interactive auto-login once
                    try:
                        if self.user_var.get().strip() and self.pass_var.get().strip():
                            cookie_try = self._get_session_cookie_for_host(host_key, port, interactive=False)
                            if cookie_try:
                                cookie_val = cookie_try
                                try:
                                    resp = requests.get(url, timeout=12, headers=headers, cookies={"ESAdmin-Cookie": cookie_val})
                                except Exception:
                                    continue
                    except Exception:
                        pass
                    if resp.status_code == 401:
                        continue

                if resp.status_code < 200 or resp.status_code >= 300:
                    continue

                try:
                    data = resp.json()
                except Exception:
                    continue

                pcts = data.get('PCTs') if isinstance(data, dict) else (data if isinstance(data, list) else None)
                if not pcts:
                    continue

                region_sum = 0
                for pct in pcts:
                    try:
                        name = pct.get('PCTName') or ''
                        cnt = self._parse_count(pct.get('PCTCnt') or 0)
                        if not name:
                            continue
                        agg_pcts[name] = agg_pcts.get(name, 0) + cnt
                        total_sum += cnt
                        region_sum += cnt
                    except Exception:
                        continue

                agg_by_region[region] = agg_by_region.get(region, 0) + region_sum
                agg_by_host[host_key] = agg_by_host.get(host_key, 0) + region_sum

        return {
            'total_sum': total_sum,
            'total_calls': total_calls,
            'agg_pcts': agg_pcts,
            'agg_by_region': agg_by_region,
            'agg_by_host': agg_by_host,
        }

    def _consultar_todos_done(self, result: dict):
        """Called on the main thread after worker completes. Updates UI and shows dialogs if needed."""
        try:
            if result is None:
                messagebox.showerror('Error', 'Error inesperado en la consulta')
                self._set_busy(False)
                return
            if 'error' in result:
                messagebox.showerror('Error', result.get('error'))
                self._set_busy(False)
                return

            total_sum = result.get('total_sum', 0)
            total_calls = result.get('total_calls', 0)
            agg_pcts = result.get('agg_pcts', {})
            agg_by_region = result.get('agg_by_region', {})

            if total_sum == 0:
                messagebox.showinfo('Resultados', 'No se obtuvieron datos de las consultas (0 ejecuciones).')
                self._set_busy(False)
                return

            self.clear_table()
            sorted_pcts = sorted(agg_pcts.items(), key=lambda x: (-x[1], x[0]))
            for name, cnt in sorted_pcts:
                self.tree.insert('', tk.END, values=(name, '', '', cnt))
            try:
                self._apply_row_tags()
            except Exception:
                pass

            summary = f"Consultar todos: {len(sorted_pcts)} filas agregadas, total ejecuciones: {total_sum} (consultas: {total_calls})"
            try:
                per_reg = ', '.join([f"{r}={v}" for r, v in agg_by_region.items()])
                summary += f" | por region: {per_reg}"
            except Exception:
                pass
            self.status_var.set(summary)
        finally:
            self._set_busy(False)


    def login(self, host: str = None, port: str = None):
        """Interactive login. If host/port provided, uses them; otherwise uses UI fields."""
        # remember whether host/port were explicitly passed as arguments
        host_arg_provided = host is not None
        port_arg_provided = port is not None
        # capture what was present in the UI before we start; we will only clear
        # host/port fields later if they were empty in the UI and we auto-selected
        # them here (to avoid erasing user selections unexpectedly).
        try:
            ui_host_before = (self.host_var.get() or '').strip()
        except Exception:
            ui_host_before = ''
        try:
            ui_port_before = (self.port_var.get() or '').strip()
        except Exception:
            ui_port_before = ''

        host = (host or ui_host_before).strip()
        port = (port or ui_port_before).strip()
        # If host/port not supplied (fields may be hidden), pick any available host/port
        if not host or not port:
            chosen_host = None
            chosen_port = None
            try:
                # prefer explicit host mapping entries
                if self.host_mapping:
                    for h, m in self.host_mapping.items():
                        chosen_host = h
                        chosen_port = str(m.get('port', self.port_var.get().strip() or '10086'))
                        break
                # fallback to presets
                if not chosen_host and getattr(self, 'host_presets', None):
                    try:
                        chosen_host = self.host_presets[0]
                        chosen_port = self.port_var.get().strip() or '10086'
                    except Exception:
                        chosen_host = None
            except Exception:
                chosen_host = None
                chosen_port = None

            if chosen_host and chosen_port:
                # if UI already provided values, preserve them; otherwise apply chosen defaults
                used_host = host or chosen_host
                used_port = port or chosen_port
                host = used_host
                port = used_port
                try:
                    self.host_var.set(host)
                    self.port_var.set(port)
                except Exception:
                    pass
                try:
                    self.status_var.set(f'Logon usando {host}:{port}')
                except Exception:
                    pass
            else:
                messagebox.showwarning("Missing", "Host y port son requeridos para logon")
                return
        cookie = self._get_session_cookie_for_host(host, port, interactive=True)
        # ensure the most-recent interactive session cookie also stored under host
        if cookie:
            self.session_cookie = cookie
            try:
                # enable UI now that user did interactive login
                self._set_ui_logged_in(True)
            except Exception:
                pass
            # If login() was called interactively (no host/port args), clear Host and Port
            # UI fields so the user must choose them manually afterwards. This makes the
            # behavior deterministic: after an explicit user logon, host/port are empty.
            try:
                if (not host_arg_provided) and (not port_arg_provided):
                    try:
                        self.host_var.set('')
                        # clear combobox display if supported
                        try:
                            if hasattr(self.host_combo, 'set'):
                                self.host_combo.set('')
                        except Exception:
                            pass
                    except Exception:
                        pass
                    try:
                        self.port_var.set('')
                    except Exception:
                        pass
            except Exception:
                pass
        

    def logoff(self):
        if not self.session_cookie or not self.logon_headers:
            messagebox.showinfo("Info", "No hay sesión activa")
            return
        host = self.host_var.get().strip()
        port = self.port_var.get().strip()
        base_url = f"http://{host}:{port}"
        try:
            microfocus_logoff(base_url, self.logon_headers, self.session_cookie)
        except Exception as e:
            messagebox.showerror("Logoff error", str(e))
            return
        self.logon_headers = None
        self.session_cookie = None
        self.status_var.set("Logoff realizado")
        # clear host/port/region fields so user must choose them manually after logoff
        try:
            try:
                self._clear_host_port_region_ui()
            except Exception:
                pass
        except Exception:
            pass
        # Recreate host/port widgets to force visible clearing (fallback)
        try:
            self._recreate_host_port_widgets()
        except Exception:
            pass
        try:
            # also clear the table like the 'Limpiar tabla' button
            try:
                self.clear_table()
            except Exception:
                pass
        except Exception:
            pass
        try:
            # disable UI again until next login
            self._set_ui_logged_in(False)
        except Exception:
            pass

    def _set_ui_logged_in(self, logged_in: bool):
        """Enable or disable the main input controls depending on login state.
        When logged_in=False, disable host/port/region controls and action buttons so the user must login first.
        """
        try:
            state = 'normal' if logged_in else 'disabled'
            # host combobox may be CTk or ttk; configure accordingly
            try:
                self.host_combo.configure(state=state)
            except Exception:
                try:
                    if state == 'disabled':
                        self.host_combo.configure(state='readonly')
                    else:
                        self.host_combo.configure(state='normal')
                except Exception:
                    pass
            try:
                self.port_entry.configure(state=state)
            except Exception:
                pass
            try:
                self.regions_combo.configure(state=state)
            except Exception:
                pass
            # mapping editor button was removed earlier; nothing to configure here
            try:
                self.btn_refresh.configure(state=state)
            except Exception:
                pass
            try:
                self.btn_clear.configure(state=state)
            except Exception:
                pass
            try:
                self.btn_consultar_todos.configure(state=state)
            except Exception:
                pass
            # show/hide host & port widgets depending on login state
            try:
                if logged_in:
                    try:
                        self.host_combo.grid()
                    except Exception:
                        pass
                    try:
                        self.port_entry.grid()
                    except Exception:
                        pass
                    # restore labels by re-gridding known positions (row=0, col=0 and col=2)
                    try:
                        inputs = self.host_combo.master
                        for child in inputs.grid_slaves(row=0, column=0):
                            try:
                                child.grid()
                            except Exception:
                                pass
                        for child in inputs.grid_slaves(row=0, column=2):
                            try:
                                child.grid()
                            except Exception:
                                pass
                    except Exception:
                        pass
                else:
                    try:
                        self.host_combo.grid_remove()
                        self.port_entry.grid_remove()
                    except Exception:
                        pass
            except Exception:
                pass
        except Exception:
            pass

    def _recreate_host_port_widgets(self):
        """Destroy and recreate host combobox and port entry to ensure the visible
        widgets are cleared. Handles both CTk and ttk variants used in _build_ui.
        """
        try:
            # determine parent frames
            parent = None
            try:
                parent = self.host_combo.master if getattr(self, 'host_combo', None) else None
            except Exception:
                parent = None
            if parent is None:
                # attempt to find a sensible parent by scanning children
                try:
                    for w in self.winfo_children():
                        if isinstance(w, tk.Frame):
                            parent = w
                            break
                except Exception:
                    parent = None

            # remove existing widgets if present
            try:
                if getattr(self, 'host_combo', None):
                    try:
                        self.host_combo.destroy()
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                if getattr(self, 'port_entry', None):
                    try:
                        self.port_entry.destroy()
                    except Exception:
                        pass
            except Exception:
                pass

            # fallback parent
            if parent is None:
                parent = self

            # create new widgets using same heuristics as _build_ui
            # factories: CTk or ttk
            EntryW = None
            ComboBox = None
            try:
                EntryW = ctk.CTkEntry if USE_CTK else ttk.Entry
            except Exception:
                EntryW = ttk.Entry
            try:
                ComboBox = ctk.CTkComboBox if (USE_CTK and hasattr(ctk, 'CTkComboBox')) else ttk.Combobox
            except Exception:
                ComboBox = ttk.Combobox

            combo_width = 220 if USE_CTK else 18
            entry_width = 80 if USE_CTK else 8

            # host combobox
            try:
                if USE_CTK and ComboBox is ctk.CTkComboBox:
                    self.host_combo = ComboBox(parent, values=getattr(self, 'host_presets', []), width=combo_width)
                    try:
                        self.host_combo.grid(row=0, column=1, sticky='w', padx=6, pady=4)
                    except Exception:
                        try:
                            self.host_combo.pack(side='left')
                        except Exception:
                            pass
                    try:
                        self.host_combo.configure(command=lambda val: (self.host_var.set(val), self._on_host_selected()))
                    except Exception:
                        pass
                else:
                    self.host_combo = ComboBox(parent, textvariable=self.host_var, values=getattr(self, 'host_presets', []), width=combo_width)
                    try:
                        self.host_combo.grid(row=0, column=1, sticky=tk.W, padx=6, pady=4)
                    except Exception:
                        try:
                            self.host_combo.pack(side='left')
                        except Exception:
                            pass
                    try:
                        self.host_combo.bind('<<ComboboxSelected>>', lambda e: self._on_host_selected())
                        self.host_combo.bind('<FocusOut>', lambda e: self._on_host_entry_changed())
                    except Exception:
                        pass
            except Exception:
                pass

            # port entry
            try:
                if USE_CTK and EntryW is ctk.CTkEntry:
                    self.port_entry = EntryW(parent, textvariable=self.port_var, width=entry_width)
                    try:
                        self.port_entry.grid(row=0, column=3, sticky='w', padx=6, pady=4)
                    except Exception:
                        try:
                            self.port_entry.pack(side='left')
                        except Exception:
                            pass
                else:
                    self.port_entry = EntryW(parent, textvariable=self.port_var, width=8)
                    try:
                        self.port_entry.grid(row=0, column=3, sticky=tk.W, padx=6, pady=4)
                    except Exception:
                        try:
                            self.port_entry.pack(side='left')
                        except Exception:
                            pass
            except Exception:
                pass

            # ensure UI refresh
            try:
                self.update_idletasks()
            except Exception:
                pass
        except Exception:
            pass

    def _clear_host_port_region_ui(self):
        """Robustly clear Host, Port and Region UI fields (StringVars and widget displays).
        Tries several strategies to support ttk and CTk widgets.
        """
        try:
            # clear bound variables first
            try:
                if hasattr(self, 'host_var'):
                    self.host_var.set('')
            except Exception:
                pass
            try:
                if hasattr(self, 'port_var'):
                    self.port_var.set('')
            except Exception:
                pass
            try:
                if hasattr(self, 'region_var'):
                    self.region_var.set('')
            except Exception:
                pass
            # also clear site/user/pass/canal vars if present
            try:
                if hasattr(self, 'canal_var'):
                    self.canal_var.set('')
            except Exception:
                pass
            try:
                if hasattr(self, 'site_var'):
                    self.site_var.set('')
            except Exception:
                pass
            try:
                if hasattr(self, 'user_var'):
                    self.user_var.set('')
            except Exception:
                pass
            try:
                if hasattr(self, 'pass_var'):
                    self.pass_var.set('')
            except Exception:
                pass
            widgets = [('host_combo', getattr(self, 'host_combo', None)),
                       ('port_entry', getattr(self, 'port_entry', None)),
                       ('regions_combo', getattr(self, 'regions_combo', None)),
                       ('site_entry', getattr(self, 'site_entry', None)),
                       ('canal_entry', getattr(self, 'canal_entry', None)),
                       ('user_entry', getattr(self, 'user_entry', None)),
                       ('pass_entry', getattr(self, 'pass_entry', None))]

            for name, w in widgets:
                if not w:
                    continue
                prev_state = None
                try:
                    # capture previous state if widget supports it
                    if hasattr(w, 'cget'):
                        try:
                            prev_state = w.cget('state')
                        except Exception:
                            prev_state = None
                except Exception:
                    prev_state = None

                # temporarily set to normal so we can clear readonly entries
                try:
                    if prev_state is not None and hasattr(w, 'configure'):
                        try:
                            w.configure(state='normal')
                        except Exception:
                            pass
                except Exception:
                    pass

                # combobox-like set
                try:
                    if hasattr(w, 'set'):
                        try:
                            w.set('')
                        except Exception:
                            pass
                except Exception:
                    pass

                # entry-like delete
                try:
                    if hasattr(w, 'delete'):
                        try:
                            w.delete(0, tk.END)
                        except Exception:
                            pass
                except Exception:
                    pass

                # clear inner children (some widgets embed an Entry)
                try:
                    for child in getattr(w, 'winfo_children', lambda: [])():
                        try:
                            if hasattr(child, 'delete'):
                                try:
                                    child.delete(0, tk.END)
                                except Exception:
                                    pass
                        except Exception:
                            pass
                except Exception:
                    pass

                # reset combobox values if supported
                try:
                    if hasattr(w, 'configure'):
                        try:
                            w.configure(values=[])
                        except Exception:
                            pass
                except Exception:
                    pass

                # restore previous state
                try:
                    if prev_state is not None and hasattr(w, 'configure'):
                        try:
                            w.configure(state=prev_state)
                        except Exception:
                            pass
                except Exception:
                    pass

            try:
                self.update_idletasks()
            except Exception:
                pass
        except Exception:
            pass

    def clear_table(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

    def _apply_row_tags(self):
        """Apply alternating row tags for zebra striping."""
        try:
            for idx, item in enumerate(self.tree.get_children()):
                tag = 'even' if idx % 2 == 0 else 'odd'
                self.tree.item(item, tags=(tag,))
        except Exception:
            pass

    def refresh(self, retry=False):
        # Read values robustly: prefer the StringVars but fallback to widget.get() when needed
        try:
            host = (self.host_var.get().strip() if getattr(self, 'host_var', None) else '').strip()
        except Exception:
            host = ''
        try:
            if not host and hasattr(self, 'host_combo') and hasattr(self.host_combo, 'get'):
                try:
                    host = str(self.host_combo.get()).strip()
                except Exception:
                    host = host
        except Exception:
            pass
        try:
            port = (self.port_var.get().strip() if getattr(self, 'port_var', None) else '').strip()
        except Exception:
            port = ''
        try:
            if not port and hasattr(self, 'port_entry') and hasattr(self.port_entry, 'get'):
                try:
                    port = str(self.port_entry.get()).strip()
                except Exception:
                    port = port
        except Exception:
            pass
        try:
            region = (self.region_var.get().strip() if getattr(self, 'region_var', None) else '').strip()
        except Exception:
            region = ''
        try:
            if not region and hasattr(self, 'regions_combo') and hasattr(self.regions_combo, 'get'):
                try:
                    region = str(self.regions_combo.get()).strip()
                except Exception:
                    region = region
        except Exception:
            pass

        # require host. If port is missing, assume the common default 10086
        # (this helps when users pick a host from the combobox but the port
        # field was cleared during login flow).
        if not host:
            messagebox.showwarning("Faltan valores", "Host es requerido")
            return
        if not port:
            port = '10086'
            try:
                self.port_var.set(port)
            except Exception:
                pass

        if not region:
            # try to infer regions for this host from mapping
            mapping = self.host_mapping.get(host)
            regions = []
            if mapping:
                regions = mapping.get('regions') or ([mapping.get('region')] if mapping.get('region') else [])
            # fallback to combobox values
            if not regions:
                try:
                    vals = list(self.regions_combo['values'])
                    regions = vals
                except Exception:
                    regions = []
            if regions:
                region = regions[0]
                try:
                    self.region_var.set(region)
                    try:
                        # CTk/ttk differences
                        self.regions_combo.set(region)
                    except Exception:
                        pass
                except Exception:
                    pass
            else:
                messagebox.showwarning("Faltan valores", "Region es requerida")
                return

        # If user selected 'All', aggregate results across all regions for this host
        if region == 'All':
            mapping = self.host_mapping.get(host)
            if not mapping:
                messagebox.showwarning('Sin mapping', 'No hay mapping para este host con regiones asociadas')
                return
            regions = mapping.get('regions') or ([mapping.get('region')] if mapping.get('region') else [])
            if not regions:
                messagebox.showwarning('Sin regiones', 'No hay regiones asociadas para este host')
                return

            params = {}

            self.status_var.set(f"Consultando todas las regiones de {host} ...")
            self.update_idletasks()

            agg_pcts = {}
            total_sum = 0
            total_calls = 0
            region_counts = {}
            failed_regions = []

            # ensure we have a session cookie for this host before iterating regions
            cookie_val = None
            try:
                cookie_val = self.session_cookies.get(host)
            except Exception:
                cookie_val = None
            if not cookie_val:
                try:
                    # try non-interactive auto-login using UI credentials
                    if self.user_var.get().strip() and self.pass_var.get().strip():
                        cookie_val = self._get_session_cookie_for_host(host, port, interactive=False)
                except Exception:
                    cookie_val = None
            # if still no cookie, attempt an automatic interactive login (no prompt)
            if not cookie_val and (self.user_var.get().strip() or self.pass_var.get().strip()):
                try:
                    self.login(host, port)
                    cookie_val = self.session_cookies.get(host)
                except Exception:
                    cookie_val = None


            for r in regions:
                url = f"http://{host}:{port}/native/v1/regions/{host}/86/{r}/active/pct"
                try:
                    headers = {"Content-Type": "application/json", "Origin": f"http://{host}:{port}", "X-Requested-With": "XMLHttpRequest"}
                    # prefer the pre-fetched cookie_val for this host; fall back to stored session_cookies
                    cookie_to_use = cookie_val or (self.session_cookies.get(host) if hasattr(self, 'session_cookies') else None)
                    cookies = {"ESAdmin-Cookie": cookie_to_use} if cookie_to_use else None
                    resp = requests.get(url, params=params, headers=headers, cookies=cookies, timeout=20)
                except Exception:
                    failed_regions.append(r)
                    continue

                total_calls += 1
                # if unauthorized, try a non-interactive auto-login first
                if resp.status_code == 401:
                    tried_auto = False
                    try:
                        auto_cookie = None
                        # try non-interactive auto-login using credentials from the form
                        if self.user_var.get().strip() and self.pass_var.get().strip():
                            auto_cookie = self._get_session_cookie_for_host(host, port, interactive=False)
                        if auto_cookie:
                            tried_auto = True
                            try:
                                resp = requests.get(url, params=params, headers=headers, cookies={"ESAdmin-Cookie": auto_cookie}, timeout=20)
                            except Exception:
                                # if retry failed, continue to next region
                                continue
                    except Exception:
                        tried_auto = False

                    # if still 401, attempt an automatic interactive login and retry once
                    if resp.status_code == 401:
                        if not retry:
                            try:
                                self.login(host, port)
                                cookie_val = self.session_cookies.get(host)
                                if cookie_val:
                                    try:
                                        resp = requests.get(url, params=params, headers=headers, cookies={"ESAdmin-Cookie": cookie_val}, timeout=20)
                                    except Exception:
                                        continue
                                else:
                                    continue
                            except Exception:
                                continue
                        else:
                            continue

                if resp.status_code < 200 or resp.status_code >= 300:
                    failed_regions.append(r)
                    continue

                try:
                    data = resp.json()
                except Exception:
                    continue

                pcts = data.get('PCTs') if isinstance(data, dict) else (data if isinstance(data, list) else None)
                if not pcts:
                    region_counts[r] = 0
                    continue

                region_sum = 0
                for pct in pcts:
                    try:
                        name = pct.get('PCTName') or ''
                        cnt_raw = pct.get('PCTCnt') or 0
                        cnt = self._parse_count(cnt_raw)
                        # aggregate by PCTName only for the 'All' view
                        key = name
                        agg_pcts[key] = agg_pcts.get(key, 0) + cnt
                        total_sum += cnt
                        region_sum += cnt
                    except Exception:
                        continue

                region_counts[r] = region_sum

            # populate tree with aggregated results
            if total_sum == 0:
                messagebox.showinfo('Resultados', 'No se obtuvieron datos de las consultas (0 ejecuciones).')
                return
            self.clear_table()
            sorted_pcts = sorted(agg_pcts.items(), key=lambda x: (-x[1], x[0]))
            for name, cnt in sorted_pcts:
                # show Group and PCTSec empty since we're aggregating across regions
                self.tree.insert('', tk.END, values=(name, '', '', cnt))
            try:
                self._apply_row_tags()
            except Exception:
                pass
            # show summary with per-region counts and any failed regions
            summary_parts = [f"All: {len(sorted_pcts)} filas, total={total_sum}, consultas={total_calls}"]
            if region_counts:
                per_reg = ", ".join([f"{r}={region_counts.get(r,0)}" for r in regions])
                summary_parts.append(f"por region: {per_reg}")
            if failed_regions:
                summary_parts.append(f"fallaron: {','.join(failed_regions)}")
            summary = ' | '.join(summary_parts)
            self.status_var.set(summary)
            # summary is shown in the status bar; avoid a blocking popup for 'All' queries
            return

        url = f"http://{host}:{port}/native/v1/regions/{host}/86/{region}/active/pct"
        params = {}

        self.status_var.set(f"Consultando {url} ...")
        self.update_idletasks()

        try:
            headers = {"Content-Type": "application/json", "Origin": f"http://{host}:{port}", "X-Requested-With": "XMLHttpRequest"}
            cookies = None
            if self.session_cookie:
                cookies = {"ESAdmin-Cookie": self.session_cookie}
            resp = requests.get(url, params=params, headers=headers, cookies=cookies, timeout=20)
            # If unauthorized, attempt login automatically and retry once
            if resp.status_code == 401:
                # try to get message from response for logging
                try:
                    data = resp.json()
                    msg = data.get('ErrorMessage') or data.get('ErrorTitle') or json.dumps(data, indent=2)
                except Exception:
                    msg = resp.text or '401 - Not authorized'

                if not retry:
                    try:
                        # attempt automatic login using current host/port so UI fields are preserved
                        self.login(host, port)
                        if self.session_cookie:
                            return self.refresh(retry=True)
                    except Exception:
                        pass

                # if retry already attempted or login didn't yield session, set status and return
                self.status_var.set('401 - No active session')
                return

            # If the server returned 404, try to parse a helpful JSON error message
            if resp.status_code == 404:
                try:
                    data_err = resp.json()
                    title = data_err.get('ErrorTitle') or f'HTTP 404'
                    msg = data_err.get('ErrorMessage') or json.dumps(data_err, indent=2)
                    messagebox.showerror(title, msg)
                    try:
                        self.status_var.set(f"404 - {title}")
                    except Exception:
                        pass
                    return
                except Exception:
                    # fallback: show raw text
                    messagebox.showerror('HTTP 404', f"El servidor respondió con: {resp.text}")
                    try:
                        self.status_var.set('404 - Not Found')
                    except Exception:
                        pass
                    return

            try:
                data = resp.json()
            except json.JSONDecodeError:
                messagebox.showerror("Respuesta inválida", f"El servidor respondió con: {resp.text}")
                self.status_var.set("JSON decode error")
                return

            if resp.status_code < 200 or resp.status_code >= 300:
                messagebox.showerror("Error HTTP", f"Código: {resp.status_code}\n{json.dumps(data, indent=2)}")
                self.status_var.set(f"Error {resp.status_code}")
                return

            self.clear_table()
            pcts = data.get('PCTs') if isinstance(data, dict) else None
            if not pcts:
                if isinstance(data, list):
                    pcts = data
                else:
                    messagebox.showinfo("Sin datos", "No se encontraron PCTs en la respuesta")
                    self.status_var.set("No data")
                    return

            for pct in pcts:
                self.tree.insert('', tk.END, values=(pct.get('PCTName'), pct.get('group'), pct.get('PCTSec'), pct.get('PCTCnt')))

            try:
                self._apply_row_tags()
            except Exception:
                pass

            self.status_var.set(f"{len(pcts)} registros")

        except requests.exceptions.Timeout:
            messagebox.showerror("Timeout", "La consulta tardó demasiado (timeout)")
            self.status_var.set("Timeout")
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Request error", str(e))
            self.status_var.set("Request error")

    


def main():
    app = MonitorApp()
    app.mainloop()


if __name__ == '__main__':
    main()
