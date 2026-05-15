#!/usr/bin/env python3
"""
MailSync — Interface graphique
Garage Martelet — garage@martelet39.fr

INSTALLATION (Mac)
==================
1. Python 3 requis : https://www.python.org/downloads/
2. pip3 install tqdm
3. python3 migration_imap_gui.py
"""

import imaplib
import email
import email.utils
import ssl
import os
import re
import sys
import time
import json
import queue
import logging
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime

# ─── Configuration par défaut ─────────────────────────────────────────────────

DEFAULT_SOURCE = {"host": "", "port": "993", "user": "", "password": ""}
DEFAULT_DESTINATION = {"host": "", "port": "993", "user": "", "password": ""}

SYSTEM_FOLDERS   = {"INBOX", "Drafts", "Sent", "Trash", "Junk", "Spam",
                    "Brouillons", "Envoyés", "Corbeille", "Indésirables",
                    "DRAFTS", "SENT", "TRASH", "JUNK", "SPAM"}
BATCH_SIZE       = 100
PAUSE_SECONDS    = 2
HEADER_BATCH     = 200

# Fichiers de données dans le dossier utilisateur (pas dans Program Files)
_USER_DIR        = os.path.join(os.path.expanduser("~"), "MailSync")
os.makedirs(_USER_DIR, exist_ok=True)
CHECKPOINT_FILE  = os.path.join(_USER_DIR, "migration_checkpoint.json")
LOG_FILE         = os.path.join(_USER_DIR, "migration_log.txt")
CONFIG_FILE      = os.path.join(_USER_DIR, "migration_profiles.json")

# ─── Palette de couleurs ──────────────────────────────────────────────────────

C_BG       = "#1c1f2b"
C_PANEL    = "#252836"
C_BORDER   = "#313550"
C_ACCENT   = "#4f7cff"
C_ACCENT2  = "#ff5f6d"
C_SUCCESS  = "#3ecf8e"
C_WARN     = "#f0a500"
C_TEXT     = "#e8eaf0"
C_MUTED    = "#7a7f9a"
C_INPUT_BG = "#1a1d2a"
C_ROW_ALT  = "#1f2233"
C_BTN_BG   = "#4f7cff"
C_BTN_FG   = "#ffffff"

import platform as _platform
_OS = _platform.system()

if _OS == "Darwin":
    _SANS = "Helvetica Neue"
    _MONO = "Menlo"
    _SZ   = 0
elif _OS == "Windows":
    _SANS = "Segoe UI"
    _MONO = "Consolas"
    _SZ   = 2
else:
    _SANS = "DejaVu Sans"
    _MONO = "DejaVu Sans Mono"
    _SZ   = 0

def _fs(n): return n + _SZ

FONT_TITLE  = (_SANS, _fs(14), "bold")
FONT_LABEL  = (_SANS, _fs(10))
FONT_SMALL  = (_SANS, _fs(10))
FONT_MONO   = (_MONO, _fs(10))
FONT_BTN    = (_SANS, _fs(10), "bold")

# ─── Gestion des profils ─────────────────────────────────────────────────────

def load_profiles():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_profiles(profiles):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(profiles, f, indent=2)

def load_config():
    profiles = load_profiles()
    last = profiles.get("__last__")
    if last and last in profiles:
        return dict(profiles[last])
    return {"name": "", "source": dict(DEFAULT_SOURCE), "destination": dict(DEFAULT_DESTINATION)}

def save_config(cfg):
    profiles = load_profiles()
    name = cfg.get("name", "").strip()
    if name:
        profiles[name] = cfg
        profiles["__last__"] = name
    save_profiles(profiles)

# ─── Checkpoint ───────────────────────────────────────────────────────────────

def load_checkpoint():
    if os.path.exists(CHECKPOINT_FILE):
        with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_checkpoint(data):
    with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

# ─── Logging vers fichier ─────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8")],
)
log = logging.getLogger(__name__)

# ─── IMAP helpers ─────────────────────────────────────────────────────────────

def connect(config):
    ctx  = ssl.create_default_context()
    conn = imaplib.IMAP4_SSL(config["host"], int(config["port"]), ssl_context=ctx)
    conn.login(config["user"], config["password"])
    return conn

def parse_folder_names(raw_list):
    names = []
    for f in raw_list:
        decoded = f.decode("utf-8", errors="replace")
        m = re.search(r'"([^"]+)"\s*$', decoded)
        name = m.group(1).strip() if m else decoded.rsplit(" ", 1)[-1].strip().strip('"')
        if name:
            names.append(name)
    return names

def ensure_folder(dst_conn, folder_name):
    try:
        dst_conn.create(f'"{folder_name}"')
    except Exception:
        pass

def count_messages(conn, folder):
    try:
        status, data = conn.select(f'"{folder}"', readonly=True)
        return int(data[0]) if status == "OK" and data[0] else 0
    except Exception:
        return 0

def list_folders(conn):
    status, folders = conn.list()
    names   = parse_folder_names(folders)
    ordered = ["INBOX"] + [n for n in names if n.upper() != "INBOX"]
    result  = []
    for name in ordered:
        count = count_messages(conn, name)
        result.append((name, count))
    return result

def scan_message_ids(conn, folder, progress_cb=None):
    result = {}
    try:
        status, data = conn.select(f'"{folder}"', readonly=True)
        if status != "OK":
            return result
        count = int(data[0]) if data[0] else 0
        if count == 0:
            return result
        status2, data2 = conn.uid("search", None, "ALL")
        if status2 != "OK" or not data2[0]:
            return result
        all_uids = data2[0].split()
        total    = len(all_uids)
        done     = 0
        for i in range(0, total, HEADER_BATCH):
            batch_uids = all_uids[i : i + HEADER_BATCH]
            uid_set    = b",".join(batch_uids)
            st, parts  = conn.uid("fetch", uid_set,
                                  "(BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)])")
            if st == "OK" and parts:
                uid_idx = 0
                for part in parts:
                    if isinstance(part, tuple) and isinstance(part[1], bytes):
                        raw = part[1].decode("utf-8", errors="replace")
                        mid = ""
                        for line in raw.splitlines():
                            if line.lower().startswith("message-id:"):
                                mid = line.split(":", 1)[1].strip()
                                break
                        if mid and uid_idx < len(batch_uids):
                            result[mid] = batch_uids[uid_idx]
                        uid_idx += 1
            done += len(batch_uids)
            if progress_cb:
                progress_cb(done, total)
    except Exception as e:
        log.warning(f"Erreur scan '{folder}' : {e}")
    return result

def fetch_message_uid(src_conn, uid):
    body = None
    internaldate = None
    status, msg_data = src_conn.uid("fetch", uid, "(RFC822 INTERNALDATE)")
    if status == "OK" and msg_data:
        for part in msg_data:
            if isinstance(part, tuple):
                try:
                    meta = part[0].decode("utf-8", errors="replace")
                    if "INTERNALDATE" in meta:
                        start    = meta.index("INTERNALDATE") + len("INTERNALDATE ")
                        raw_date = meta[start:].split('"')[1]
                        internaldate = imaplib.Internaldate2tuple(
                            f'* 1 FETCH (INTERNALDATE "{raw_date}")'.encode()
                        )
                except Exception:
                    pass
                if isinstance(part[1], bytes) and len(part[1]) > 0:
                    body = part[1]
    if body is None:
        status2, msg_data2 = src_conn.uid("fetch", uid, "BODY[]")
        if status2 == "OK" and msg_data2 and isinstance(msg_data2[0], tuple):
            if isinstance(msg_data2[0][1], bytes):
                body = msg_data2[0][1]
    if internaldate is None and body is not None:
        try:
            msg      = email.message_from_bytes(body)
            date_str = msg.get("Date")
            if date_str:
                parsed = email.utils.parsedate_tz(date_str)
                if parsed:
                    internaldate = time.gmtime(email.utils.mktime_tz(parsed))
        except Exception:
            pass
    return body, internaldate

def purge_infomaniak(conn, log_cb):
    log_cb("Purge Infomaniak en cours...")
    status, folders = conn.list()
    if status != "OK":
        log_cb("  Impossible de lister les dossiers.", "error")
        return
    folder_names = parse_folder_names(folders)
    ordered = ["INBOX"] + [n for n in folder_names if n.upper() != "INBOX"]
    for name in ordered:
        try:
            status, data = conn.select(f'"{name}"')
            if status != "OK":
                continue
            count = int(data[0]) if data[0] else 0
            if count == 0:
                continue
            conn.store("1:*", "+FLAGS", "\\Deleted")
            conn.expunge()
            log_cb(f"  {name} : {count} message(s) supprimé(s).")
        except Exception as e:
            log_cb(f"  Erreur sur '{name}' : {e}", "warn")
    for name in folder_names:
        if name.upper() in {s.upper() for s in SYSTEM_FOLDERS}:
            continue
        try:
            conn.delete(f'"{name}"')
        except Exception:
            pass
    log_cb("Purge terminée.", "success")

# ─── Application principale ───────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        if _OS == "Windows":
            try:
                import ctypes
                # PerMonitorV2 awareness (Windows 10 1703+)
                ctypes.windll.shcore.SetProcessDpiAwareness(2)
            except Exception:
                try:
                    ctypes.windll.user32.SetProcessDPIAware()
                except Exception:
                    pass
        self.title("MailSync")
        self.configure(bg=C_BG)
        self.resizable(True, True)
        w, h = (1020, 800) if _OS == "Windows" else (980, 760)
        self.geometry(f"{w}x{h}")
        self.minsize(860, 640)

        self.cfg       = load_config()
        self.msg_queue = queue.Queue()
        self.running   = False
        self.src_conn  = None
        self.dst_conn  = None
        self.folder_rows = []    # [(name, src_count, dst_count, delta, BoolVar)]

        self._build_ui()
        self.after(100, self._update_dynamic_labels)
        self._poll_queue()

    # ── Construction de l'interface ───────────────────────────────────────────

    def _build_ui(self):
        # Barre de titre
        header = tk.Frame(self, bg=C_BG)
        header.pack(fill="x", padx=24, pady=(20, 0))
        tk.Label(header, text="MailSync", font=FONT_TITLE,
                 bg=C_BG, fg=C_TEXT).pack(side="left")
        tk.Label(header, text="Paritel  →  Infomaniak",
                 font=FONT_SMALL, bg=C_BG, fg=C_MUTED).pack(side="left", padx=16)

        # Notebook (onglets)
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure("TNotebook",
                        background=C_BG, borderwidth=0, tabmargins=0)
        style.configure("TNotebook.Tab",
                        background=C_PANEL, foreground=C_MUTED,
                        padding=[18, 8], font=FONT_SMALL, borderwidth=0)
        style.map("TNotebook.Tab",
                  background=[("selected", C_BG)],
                  foreground=[("selected", C_TEXT)])

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=16, pady=12)

        tab_conf  = tk.Frame(nb, bg=C_BG)
        tab_mig   = tk.Frame(nb, bg=C_BG)
        nb.add(tab_conf, text="  Configuration  ")
        nb.add(tab_mig,  text="  Migration  ")

        self._build_config_tab(tab_conf)
        self._build_migration_tab(tab_mig)

    def _section(self, parent, title):
        fr = tk.LabelFrame(parent, text=f"  {title}  ",
                           bg=C_PANEL, fg=C_MUTED,
                           font=FONT_SMALL,
                           bd=1, relief="flat",
                           highlightbackground=C_BORDER,
                           highlightthickness=1)
        return fr

    def _field(self, parent, row, label, value, show=""):
        tk.Label(parent, text=label, bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL, anchor="w").grid(
            row=row, column=0, sticky="w", padx=(14, 8), pady=6)
        var = tk.StringVar(value=value)
        e = tk.Entry(parent, textvariable=var, show=show,
                     bg=C_INPUT_BG, fg=C_TEXT, font=FONT_SMALL,
                     insertbackground=C_TEXT, relief="flat",
                     highlightbackground=C_BORDER, highlightthickness=1,
                     width=36)
        e.grid(row=row, column=1, sticky="ew", padx=(0, 14), pady=6)
        return var

    def _build_config_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        # Source
        self.fr_src_label = tk.StringVar(value="Serveur source")
        fr_src = self._section(parent, "Serveur source")
        fr_src.grid(row=0, column=0, sticky="nsew", padx=(8, 6), pady=8)
        fr_src.columnconfigure(1, weight=1)

        self.src_host = self._field(fr_src, 0, "Hôte IMAP",  self.cfg["source"]["host"])
        self.src_port = self._field(fr_src, 1, "Port",       self.cfg["source"]["port"])
        self.src_user = self._field(fr_src, 2, "Utilisateur", self.cfg["source"]["user"])
        self.src_pass = self._field(fr_src, 3, "Mot de passe", self.cfg["source"]["password"], show="●")

        btn_test_src = tk.Button(fr_src, text="Tester la connexion",
                                 command=lambda: self._test_connect("source"),
                                 bg=C_PANEL, fg=C_ACCENT, font=FONT_SMALL,
                                 relief="flat", cursor="hand2",
                                 highlightbackground=C_BORDER, highlightthickness=1,
                                 padx=10, pady=4)
        btn_test_src.grid(row=4, column=0, columnspan=2,
                          sticky="w", padx=14, pady=(4, 12))

        # Destination
        fr_dst = self._section(parent, "Serveur destination")
        fr_dst.grid(row=0, column=1, sticky="nsew", padx=(6, 8), pady=8)
        fr_dst.columnconfigure(1, weight=1)

        self.dst_host = self._field(fr_dst, 0, "Hôte IMAP",  self.cfg["destination"]["host"])
        self.dst_port = self._field(fr_dst, 1, "Port",       self.cfg["destination"]["port"])
        self.dst_user = self._field(fr_dst, 2, "Utilisateur", self.cfg["destination"]["user"])
        self.dst_pass = self._field(fr_dst, 3, "Mot de passe", self.cfg["destination"]["password"], show="●")

        btn_test_dst = tk.Button(fr_dst, text="Tester la connexion",
                                 command=lambda: self._test_connect("destination"),
                                 bg=C_PANEL, fg=C_ACCENT, font=FONT_SMALL,
                                 relief="flat", cursor="hand2",
                                 highlightbackground=C_BORDER, highlightthickness=1,
                                 padx=10, pady=4)
        btn_test_dst.grid(row=4, column=0, columnspan=2,
                          sticky="w", padx=14, pady=(4, 12))

        # Gestion des profils
        fr_prof = self._section(parent, "Profils de configuration")
        fr_prof.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        fr_prof.columnconfigure(1, weight=1)

        tk.Label(fr_prof, text="Nom du profil :", bg=C_PANEL, fg=C_MUTED,
                 font=FONT_SMALL).grid(row=0, column=0, padx=(14, 8), pady=8, sticky="w")
        self.profile_name = tk.StringVar(value=self.cfg.get("name", ""))
        tk.Entry(fr_prof, textvariable=self.profile_name,
                 bg=C_INPUT_BG, fg=C_TEXT, font=FONT_SMALL,
                 insertbackground=C_TEXT, relief="flat",
                 highlightbackground=C_BORDER, highlightthickness=1,
                 width=28).grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=8)

        fr_btns = tk.Frame(fr_prof, bg=C_PANEL)
        fr_btns.grid(row=0, column=2, padx=(0, 14), pady=8)
        self._btn(fr_btns, "Sauvegarder", self._save_config, color=C_SUCCESS).pack(side="left", padx=(0, 6))
        self._btn(fr_btns, "Charger…", self._load_profile_dialog, color=C_ACCENT).pack(side="left", padx=(0, 6))
        self._btn(fr_btns, "Supprimer", self._delete_profile, color=C_ACCENT2).pack(side="left")

        # Zone de log de connexion
        fr_log = self._section(parent, "Log de connexion")
        fr_log.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=8, pady=(0, 8))
        parent.rowconfigure(2, weight=1)
        self.conn_log = scrolledtext.ScrolledText(
            fr_log, bg=C_INPUT_BG, fg=C_TEXT, font=FONT_MONO,
            relief="flat", height=6, state="disabled",
            wrap="word", insertbackground=C_TEXT)
        self.conn_log.pack(fill="both", expand=True, padx=8, pady=8)
        self._tag_colors(self.conn_log)

    def _build_migration_tab(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)
        parent.rowconfigure(3, weight=2)

        # Barre d'actions
        fr_actions = tk.Frame(parent, bg=C_BG)
        fr_actions.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))

        self.btn_analyse = self._btn(fr_actions, "Analyser les dossiers",
                                     self._start_analyse, color=C_ACCENT)
        self.btn_analyse.pack(side="left", padx=(0, 8))

        self.btn_select_all = self._btn(fr_actions, "Tout sélectionner",
                                        self._select_all, color=C_PANEL)
        self.btn_select_all.pack(side="left", padx=(0, 4))

        self.btn_select_none = self._btn(fr_actions, "Tout désélectionner",
                                         self._select_none, color=C_PANEL)
        self.btn_select_none.pack(side="left", padx=(0, 16))

        self.btn_migrate = self._btn(fr_actions, "Lancer la migration",
                                     self._start_migration, color=C_ACCENT2)
        self.btn_migrate.pack(side="left", padx=(0, 8))
        self.btn_migrate.config(state="disabled")

        self.btn_purge = self._btn(fr_actions, "⚠  Purger Infomaniak",
                                   self._confirm_purge, color=C_WARN)
        self.btn_purge.pack(side="right")

        # Tableau des dossiers
        fr_table = self._section(parent, "Dossiers")
        fr_table.grid(row=1, column=0, sticky="nsew", padx=8, pady=4)
        fr_table.columnconfigure(0, weight=1)
        fr_table.rowconfigure(0, weight=1)

        cols = ("sel", "folder", "src", "dst", "delta")
        self.tree = ttk.Treeview(fr_table, columns=cols,
                                 show="headings", selectmode="none")

        style = ttk.Style()
        style.configure("Treeview",
                        background=C_PANEL, fieldbackground=C_PANEL,
                        foreground=C_TEXT, rowheight=26,
                        font=FONT_SMALL, borderwidth=0)
        style.configure("Treeview.Heading",
                        background=C_INPUT_BG, foreground=C_MUTED,
                        font=FONT_SMALL, relief="flat")
        style.map("Treeview", background=[("selected", C_BORDER)])

        self.tree.heading("sel",    text="")
        self.tree.heading("folder", text="Dossier")
        self.tree.heading("src",    text="Paritel")
        self.tree.heading("dst",    text="Infomaniak")
        self.tree.heading("delta",  text="À copier")

        self.tree.column("sel",    width=36,  stretch=False, anchor="center")
        self.tree.column("folder", width=260, stretch=True,  anchor="w")
        self.tree.column("src",    width=90,  stretch=False, anchor="center")
        self.tree.column("dst",    width=100, stretch=False, anchor="center")
        self.tree.column("delta",  width=90,  stretch=False, anchor="center")

        self.tree.tag_configure("has_delta", foreground=C_ACCENT)
        self.tree.tag_configure("done",      foreground=C_SUCCESS)
        self.tree.tag_configure("alt",       background=C_ROW_ALT)

        vsb = ttk.Scrollbar(fr_table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        self.tree.bind("<Button-1>", self._on_tree_click)

        # Barre de progression globale
        fr_prog = tk.Frame(parent, bg=C_BG)
        fr_prog.grid(row=2, column=0, sticky="ew", padx=8, pady=(6, 2))
        fr_prog.columnconfigure(1, weight=1)

        tk.Label(fr_prog, text="Progression :", bg=C_BG,
                 fg=C_MUTED, font=FONT_SMALL).grid(row=0, column=0, padx=(0, 8))

        style.configure("Green.Horizontal.TProgressbar",
                        troughcolor=C_INPUT_BG, background=C_ACCENT,
                        bordercolor=C_BORDER, lightcolor=C_ACCENT,
                        darkcolor=C_ACCENT)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(fr_prog, variable=self.progress_var,
                                            maximum=100, length=400,
                                            style="Green.Horizontal.TProgressbar")
        self.progress_bar.grid(row=0, column=1, sticky="ew")

        self.progress_lbl = tk.Label(fr_prog, text="", bg=C_BG,
                                     fg=C_MUTED, font=FONT_SMALL)
        self.progress_lbl.grid(row=0, column=2, padx=(10, 0))

        # Log de migration
        fr_log = self._section(parent, "Journal")
        fr_log.grid(row=3, column=0, sticky="nsew", padx=8, pady=(4, 8))
        fr_log.columnconfigure(0, weight=1)
        fr_log.rowconfigure(0, weight=1)

        self.mig_log = scrolledtext.ScrolledText(
            fr_log, bg=C_INPUT_BG, fg=C_TEXT, font=FONT_MONO,
            relief="flat", state="disabled", wrap="word",
            insertbackground=C_TEXT)
        self.mig_log.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self._tag_colors(self.mig_log)

    # ── Widgets utilitaires ───────────────────────────────────────────────────

    def _btn(self, parent, text, command, color=C_BTN_BG):
        fg = C_BG if color in (C_SUCCESS, C_WARN, C_ACCENT, C_ACCENT2) else C_TEXT
        if color == C_PANEL:
            fg = C_TEXT
        b = tk.Button(parent, text=text, command=command,
                      bg=color, fg=fg, font=FONT_BTN,
                      relief="flat", cursor="hand2",
                      padx=14, pady=6, bd=0,
                      activebackground=color, activeforeground=fg)
        return b

    def _tag_colors(self, widget):
        widget.tag_config("info",    foreground=C_TEXT)
        widget.tag_config("success", foreground=C_SUCCESS)
        widget.tag_config("warn",    foreground=C_WARN)
        widget.tag_config("error",   foreground=C_ACCENT2)
        widget.tag_config("muted",   foreground=C_MUTED)

    # ── Polling queue (thread-safe UI updates) ────────────────────────────────

    def _poll_queue(self):
        try:
            while True:
                item = self.msg_queue.get_nowait()
                kind = item.get("kind")

                if kind == "log":
                    self._append_log(self.mig_log,
                                     item["text"], item.get("tag", "info"))

                elif kind == "conn_log":
                    self._append_log(self.conn_log,
                                     item["text"], item.get("tag", "info"))

                elif kind == "progress":
                    val = item.get("value", 0)
                    self.progress_var.set(val)
                    self.progress_lbl.config(text=item.get("label", ""))

                elif kind == "folders_ready":
                    self._populate_table(item["rows"])
                    self.btn_migrate.config(state="normal")
                    self.btn_analyse.config(state="normal", text="Analyser les dossiers")
                    self.running = False

                elif kind == "migration_done":
                    self.btn_migrate.config(state="normal",
                                            text="Lancer la migration")
                    self.btn_analyse.config(state="normal")
                    self.btn_purge.config(state="normal")
                    self.running = False

                elif kind == "error_modal":
                    messagebox.showerror("Erreur", item["text"])
                    self.btn_analyse.config(state="normal",
                                            text="Analyser les dossiers")
                    self.btn_migrate.config(state="normal")
                    self.btn_purge.config(state="normal")
                    self.running = False

                elif kind == "update_row":
                    iid   = item["iid"]
                    delta = item["delta"]
                    sel   = "☑" if item["checked"] else "☐"
                    tag   = "done" if delta == 0 else "has_delta"
                    self.tree.item(iid, values=(
                        sel, item["folder"], item["src"],
                        item["dst"], delta if delta > 0 else "✓"
                    ), tags=(tag,))

        except queue.Empty:
            pass
        self.after(60, self._poll_queue)

    def _append_log(self, widget, text, tag="info"):
        widget.config(state="normal")
        widget.insert("end", text + "\n", tag)
        widget.see("end")
        widget.config(state="disabled")

    # ── Config ────────────────────────────────────────────────────────────────

    def _get_src_config(self):
        return {
            "host":     self.src_host.get().strip(),
            "port":     self.src_port.get().strip(),
            "user":     self.src_user.get().strip(),
            "password": self.src_pass.get(),
        }

    def _get_dst_config(self):
        return {
            "host":     self.dst_host.get().strip(),
            "port":     self.dst_port.get().strip(),
            "user":     self.dst_user.get().strip(),
            "password": self.dst_pass.get(),
        }

    def _save_config(self):
        name = self.profile_name.get().strip()
        if not name:
            messagebox.showwarning("Nom requis", "Saisissez un nom de profil avant de sauvegarder.")
            return
        self.cfg = {"name": name,
                    "source": self._get_src_config(),
                    "destination": self._get_dst_config()}
        save_config(self.cfg)
        self._update_dynamic_labels()
        self._append_log(self.conn_log, f"Profil \"{name}\" sauvegardé.", "success")

    def _load_profile_dialog(self):
        profiles = {k: v for k, v in load_profiles().items()
                    if not k.startswith("__")}
        if not profiles:
            messagebox.showinfo("Aucun profil", "Aucun profil sauvegardé.")
            return
        win = tk.Toplevel(self)
        win.title("Charger un profil")
        win.configure(bg=C_BG)
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="Sélectionner un profil :", bg=C_BG,
                 fg=C_TEXT, font=FONT_SMALL).pack(padx=20, pady=(16, 8))
        listbox = tk.Listbox(win, bg=C_INPUT_BG, fg=C_TEXT, font=FONT_SMALL,
                             selectbackground=C_ACCENT, relief="flat",
                             highlightbackground=C_BORDER, highlightthickness=1,
                             width=40, height=min(len(profiles), 10))
        for name in sorted(profiles.keys()):
            listbox.insert("end", name)
        listbox.pack(padx=20, pady=(0, 12))
        def _load():
            sel = listbox.curselection()
            if not sel:
                return
            name = listbox.get(sel[0])
            cfg  = profiles[name]
            self.cfg = cfg
            self.profile_name.set(cfg.get("name", name))
            self.src_host.set(cfg["source"]["host"])
            self.src_port.set(cfg["source"]["port"])
            self.src_user.set(cfg["source"]["user"])
            self.src_pass.set(cfg["source"]["password"])
            self.dst_host.set(cfg["destination"]["host"])
            self.dst_port.set(cfg["destination"]["port"])
            self.dst_user.set(cfg["destination"]["user"])
            self.dst_pass.set(cfg["destination"]["password"])
            self._update_dynamic_labels()
            profiles2 = load_profiles()
            profiles2["__last__"] = name
            save_profiles(profiles2)
            self._append_log(self.conn_log, f"Profil \"{name}\" chargé.", "success")
            win.destroy()
        self._btn(win, "Charger", _load, color=C_ACCENT).pack(pady=(0, 16))

    def _delete_profile(self):
        name = self.profile_name.get().strip()
        if not name:
            return
        profiles = load_profiles()
        if name not in profiles:
            messagebox.showinfo("Profil introuvable", f"Le profil \"{name}\" n'existe pas.")
            return
        if not messagebox.askyesno("Supprimer", f"Supprimer le profil \"{name}\" ?"):
            return
        del profiles[name]
        if profiles.get("__last__") == name:
            del profiles["__last__"]
        save_profiles(profiles)
        self._append_log(self.conn_log, f"Profil \"{name}\" supprimé.", "warn")

    def _update_dynamic_labels(self):
        if not hasattr(self, "lbl_subtitle"):
            return
        src_host  = self.src_host.get().strip()
        dst_host  = self.dst_host.get().strip()
        src_label = src_host if src_host else "Source"
        dst_label = dst_host if dst_host else "Destination"
        name      = self.profile_name.get().strip()
        subtitle  = f"{src_label}  →  {dst_label}"
        if name:
            subtitle = f"{name}  —  {subtitle}"
        self.lbl_subtitle.config(text=subtitle)
        # Bouton purge dynamique
        self.btn_purge.config(text=f"⚠  Purger {dst_label}")
        # En-têtes tableau
        self.tree.heading("src", text=src_label)
        self.tree.heading("dst", text=dst_label)

    def _test_connect(self, side):
        cfg = self._get_src_config() if side == "source" else self._get_dst_config()
        label = (self.src_host.get().strip() or "Source") if side == "source" else (self.dst_host.get().strip() or "Destination")
        self._append_log(self.conn_log, f"Test de connexion à {label}...", "muted")
        def _run():
            try:
                c = connect(cfg)
                c.logout()
                self.msg_queue.put({"kind": "conn_log",
                                    "text": f"  ✓ Connexion à {label} réussie.",
                                    "tag": "success"})
            except Exception as e:
                self.msg_queue.put({"kind": "conn_log",
                                    "text": f"  ✗ Échec : {e}",
                                    "tag": "error"})
        threading.Thread(target=_run, daemon=True).start()

    # ── Tableau ───────────────────────────────────────────────────────────────

    def _populate_table(self, rows):
        self.tree.delete(*self.tree.get_children())
        self.folder_rows = []
        for i, (name, src, dst, delta) in enumerate(rows):
            checked = delta > 0
            var     = tk.BooleanVar(value=checked)
            sel     = "☑" if checked else "☐"
            tag     = "has_delta" if delta > 0 else "done"
            if i % 2 == 1:
                tag = (tag, "alt")
            iid = self.tree.insert("", "end",
                                   values=(sel, name, src, dst,
                                           delta if delta > 0 else "✓"),
                                   tags=(tag,))
            self.folder_rows.append((iid, name, src, dst, delta, var))

    def _on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        for idx, (row_iid, name, src, dst, delta, var) in enumerate(self.folder_rows):
            if row_iid == iid:
                var.set(not var.get())
                sel = "☑" if var.get() else "☐"
                current = list(self.tree.item(iid, "values"))
                current[0] = sel
                self.tree.item(iid, values=current)
                break

    def _select_all(self):
        for iid, name, src, dst, delta, var in self.folder_rows:
            var.set(True)
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "☑"
            self.tree.item(iid, values=vals)

    def _select_none(self):
        for iid, name, src, dst, delta, var in self.folder_rows:
            var.set(False)
            vals = list(self.tree.item(iid, "values"))
            vals[0] = "☐"
            self.tree.item(iid, values=vals)

    # ── Analyse ───────────────────────────────────────────────────────────────

    def _start_analyse(self):
        if self.running:
            return
        self.running = True
        self.btn_analyse.config(state="disabled", text="Analyse en cours...")
        self.btn_migrate.config(state="disabled")
        threading.Thread(target=self._analyse_thread, daemon=True).start()

    def _analyse_thread(self):
        q = self.msg_queue
        q.put({"kind": "log", "text": "Connexion aux serveurs...", "tag": "muted"})
        try:
            src = connect(self._get_src_config())
            dst = connect(self._get_dst_config())
        except Exception as e:
            q.put({"kind": "error_modal", "text": f"Connexion impossible :\n{e}"})
            return

        q.put({"kind": "log", "text": "Récupération des dossiers...", "tag": "muted"})
        all_folders = list_folders(src)
        rows        = []

        total = len(all_folders)
        for i, (name, src_count) in enumerate(all_folders):
            dst_count = count_messages(dst, name)
            delta     = max(0, src_count - dst_count)
            rows.append((name, src_count, dst_count, delta))
            pct = int((i + 1) / total * 100)
            q.put({"kind": "progress",
                   "value": pct,
                   "label": f"{i+1}/{total} dossiers"})

        try:
            src.logout()
            dst.logout()
        except Exception:
            pass

        total_delta = sum(d for _, _, _, d in rows)
        q.put({"kind": "log",
               "text": f"Analyse terminée — {total} dossiers, {total_delta} message(s) à copier.",
               "tag": "success"})
        q.put({"kind": "progress", "value": 100, "label": "Analyse terminée"})
        q.put({"kind": "folders_ready", "rows": rows})

    # ── Migration ─────────────────────────────────────────────────────────────

    def _start_migration(self):
        if self.running:
            return
        selected = [name for _, name, _, _, _, var in self.folder_rows if var.get()]
        if not selected:
            messagebox.showwarning("Aucun dossier",
                                   "Sélectionnez au moins un dossier à migrer.")
            return
        self.running = True
        self.btn_migrate.config(state="disabled", text="Migration en cours...")
        self.btn_analyse.config(state="disabled")
        self.btn_purge.config(state="disabled")
        threading.Thread(target=self._migration_thread,
                         args=(selected,), daemon=True).start()

    def _migration_thread(self, folders):
        q          = self.msg_queue
        checkpoint = load_checkpoint()

        q.put({"kind": "log", "text": "="*52, "tag": "muted"})
        q.put({"kind": "log",
               "text": f"Migration démarrée — {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
               "tag": "info"})
        q.put({"kind": "log", "text": "="*52, "tag": "muted"})

        try:
            src = connect(self._get_src_config())
            dst = connect(self._get_dst_config())
        except Exception as e:
            q.put({"kind": "error_modal", "text": f"Connexion impossible :\n{e}"})
            return

        total_copied = 0
        n_folders    = len(folders)

        for fi, folder in enumerate(folders):
            q.put({"kind": "log", "text": f"\nDossier : {folder}", "tag": "info"})
            q.put({"kind": "progress",
                   "value": int(fi / n_folders * 100),
                   "label": f"Dossier {fi+1}/{n_folders}"})

            try:
                # Scan source
                q.put({"kind": "log", "text": "  Scan Paritel...", "tag": "muted"})
                def _prog_src(done, total, f=folder):
                    q.put({"kind": "log",
                           "text": f"  Paritel : {done}/{total} en-têtes lus...",
                           "tag": "muted"}) if done == total else None
                src_ids = scan_message_ids(src, folder, progress_cb=_prog_src)
                q.put({"kind": "log",
                       "text": f"  {len(src_ids)} messages sur Paritel.", "tag": "info"})

                if not src_ids:
                    q.put({"kind": "log", "text": "  Dossier vide, ignoré.", "tag": "muted"})
                    continue

                # Scan destination
                ck_key     = f"dst_msgids__{folder}"
                cached_dst = set(checkpoint.get(ck_key, []))
                if cached_dst:
                    dst_ids = cached_dst
                    q.put({"kind": "log",
                           "text": f"  {len(dst_ids)} Message-ID en cache (checkpoint).",
                           "tag": "muted"})
                else:
                    q.put({"kind": "log", "text": "  Scan Infomaniak...", "tag": "muted"})
                    ensure_folder(dst, folder)
                    dst_map = scan_message_ids(dst, folder)
                    dst_ids = set(dst_map.keys())
                    q.put({"kind": "log",
                           "text": f"  {len(dst_ids)} messages sur Infomaniak.", "tag": "info"})
                    checkpoint[ck_key] = list(dst_ids)
                    save_checkpoint(checkpoint)

                # Diff
                to_copy = {mid: uid for mid, uid in src_ids.items()
                           if mid not in dst_ids}
                q.put({"kind": "log",
                       "text": f"  À copier : {len(to_copy)} message(s).",
                       "tag": "success" if len(to_copy) > 0 else "muted"})

                if not to_copy:
                    q.put({"kind": "log",
                           "text": "  Infomaniak est déjà à jour.", "tag": "success"})
                    # Mise à jour du tableau
                    for row in self.folder_rows:
                        if row[1] == folder:
                            q.put({"kind": "update_row",
                                   "iid": row[0], "folder": folder,
                                   "src": row[2], "dst": row[3],
                                   "delta": 0, "checked": False})
                    continue

                ensure_folder(dst, folder)
                pending = list(to_copy.values())
                copied  = 0
                errors  = 0

                for i, uid in enumerate(pending):
                    pct_folder = int((fi + (i + 1) / len(pending)) / n_folders * 100)
                    q.put({"kind": "progress",
                           "value": pct_folder,
                           "label": f"{folder} — {i+1}/{len(pending)}"})
                    try:
                        raw_email, internaldate = fetch_message_uid(src, uid)
                        if raw_email is None:
                            raise Exception("Corps introuvable")
                        dst.append(f'"{folder}"', None, internaldate, raw_email)
                        copied += 1
                        if copied % BATCH_SIZE == 0:
                            time.sleep(PAUSE_SECONDS)
                    except Exception as e:
                        errors += 1
                        q.put({"kind": "log",
                               "text": f"  ⚠ Erreur UID {uid} : {e}", "tag": "warn"})
                        if errors > 50:
                            q.put({"kind": "log",
                                   "text": "  Trop d'erreurs, dossier abandonné.", "tag": "error"})
                            break

                if ck_key in checkpoint:
                    del checkpoint[ck_key]
                    save_checkpoint(checkpoint)

                total_copied += copied
                msg = f"  → {copied} message(s) copié(s)"
                if errors:
                    msg += f", {errors} erreur(s)"
                q.put({"kind": "log", "text": msg,
                       "tag": "success" if not errors else "warn"})
                log.info(msg)

                # Mise à jour du tableau
                for row in self.folder_rows:
                    if row[1] == folder:
                        new_delta = max(0, row[4] - copied)
                        q.put({"kind": "update_row",
                               "iid": row[0], "folder": folder,
                               "src": row[2], "dst": row[3] + copied,
                               "delta": new_delta, "checked": False})

            except Exception as e:
                q.put({"kind": "log",
                       "text": f"Erreur inattendue sur '{folder}' : {e}", "tag": "error"})
                log.error(f"Erreur inattendue sur '{folder}' : {e}")
            finally:
                try:
                    src.noop()
                except Exception:
                    try:
                        src = connect(self._get_src_config())
                    except Exception:
                        pass
                try:
                    dst.noop()
                except Exception:
                    try:
                        dst = connect(self._get_dst_config())
                    except Exception:
                        pass

        try:
            src.logout()
            dst.logout()
        except Exception:
            pass

        q.put({"kind": "log", "text": "\n" + "="*52, "tag": "muted"})
        q.put({"kind": "log",
               "text": f"Migration terminée. Total copié : {total_copied} message(s).",
               "tag": "success"})
        q.put({"kind": "log", "text": "="*52, "tag": "muted"})
        q.put({"kind": "progress", "value": 100, "label": "Terminé"})
        q.put({"kind": "migration_done"})

    # ── Purge ─────────────────────────────────────────────────────────────────

    def _confirm_purge(self):
        dst_label = self.dst_host.get().strip() or "Destination"
        if not messagebox.askyesno(
            f"Purge {dst_label}",
            f"Supprimer TOUS les messages et dossiers de la boîte {dst_label} ?\n\n"
            "⚠  Cette action est IRRÉVERSIBLE.",
            icon="warning"
        ):
            return
        if not messagebox.askyesno(
            "Dernière confirmation",
            f"Confirmer la suppression définitive de toute la boîte {dst_label} ?",
            icon="warning"
        ):
            return
        self.running = True
        self.btn_purge.config(state="disabled")
        self.btn_migrate.config(state="disabled")
        self.btn_analyse.config(state="disabled")
        threading.Thread(target=self._purge_thread, daemon=True).start()

    def _purge_thread(self):
        q = self.msg_queue
        try:
            dst = connect(self._get_dst_config())
        except Exception as e:
            q.put({"kind": "error_modal", "text": f"Connexion impossible :\n{e}"})
            return

        def _log(text, tag="info"):
            q.put({"kind": "log", "text": text, "tag": tag})

        purge_infomaniak(dst, _log)
        try:
            dst.logout()
        except Exception:
            pass

        if os.path.exists(CHECKPOINT_FILE):
            os.remove(CHECKPOINT_FILE)
        _log("Checkpoint réinitialisé.", "muted")

        q.put({"kind": "migration_done"})

# ─── Lancement ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
