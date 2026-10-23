"""
SpeziSmith - Grafische Oberfläche

Modernes, schlichtes UI (Sun-Valley-Theme) mit:
- Sprachauswahl Deutsch/English (steuert die analysierte Vorlage)
- allen Feldern der Spezifikation als wählbare Zeilen mit Eingabekästchen
- Drag & Drop für beliebig viele Spezifikationen gleichzeitig
- Ausgabe in Zielordner oder per "Speichern unter"-Dialog
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:  # Fallback ohne Drag & Drop
    DND_AVAILABLE = False

import sv_ttk

from . import APP_NAME, __version__
from .engine import (
    Rule,
    analyze_template,
    apply_to_file,
    detect_language,
    unique_target_path,
)
from .resources import (
    LOGO_PNG,
    TEMPLATE_DE,
    TEMPLATE_EN,
    load_config,
    resource_path,
    save_config,
)

NAVY = "#1F2942"
ORANGE = "#F08C21"

LANG_DE = "de"
LANG_EN = "en"
LANG_LABELS = {LANG_DE: "Deutsch", LANG_EN: "English"}

FILE_TYPES = [("Excel-Spezifikation", "*.xlsx *.xlsm")]


class ToolTip:
    """Kleiner Hover-Tooltip für ttk-Widgets."""

    def __init__(self, widget: tk.Widget, text: str):
        self.text = text
        self.tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _event=None):
        if self.tip is not None:
            return
        self.tip = tk.Toplevel()
        self.tip.overrideredirect(True)
        self.tip.attributes("-topmost", True)
        label = tk.Label(
            self.tip, text=self.text, background="#FFFFE0", relief="solid",
            borderwidth=1, padx=6, pady=3, wraplength=320, justify="left",
        )
        label.pack()
        x = self.tip.winfo_pointerx() + 12
        y = self.tip.winfo_pointery() + 16
        self.tip.geometry(f"+{x}+{y}")

    def _hide(self, _event=None):
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


class FieldRow:
    """Eine Zeile in der Feldliste (Checkbox | Zelle | Feld | Wert | Eingabe)."""

    def __init__(self, app: "SpeziSmithApp", parent: tk.Widget, field):
        self.app = app
        self.field = field
        self.checked = tk.BooleanVar(value=False)
        self.new_value = tk.StringVar()
        self._traced = True

        self.cb = ttk.Checkbutton(parent, variable=self.checked, command=app.refresh_counter)
        self.l_coord = ttk.Label(parent, text=field.coord, width=7,
                                 font=("Consolas", 10, "bold"))
        self.l_label = ttk.Label(parent, text=field.label or "—", width=34,
                                 anchor="w")
        value_text = field.display_value
        if len(value_text) > 60:
            value_text = value_text[:60] + "…"
        self.l_value = ttk.Label(parent, text=value_text, width=44, anchor="w",
                                 foreground="#666666")
        self.entry = ttk.Entry(parent, textvariable=self.new_value, width=36)

        self.widgets = (self.cb, self.l_coord, self.l_label, self.l_value, self.entry)
        self.new_value.trace_add("write", self._on_entry)
        tooltip_text = f"{field.coord}: {field.display_value}"
        for w in (self.l_label, self.l_value):
            ToolTip(w, tooltip_text)

    def _on_entry(self, *_):
        if not self._traced:
            return
        text = self.new_value.get()
        if text and not self.checked.get():
            self.checked.set(True)
        elif not text:
            self.checked.set(False)
        self.app.refresh_counter()

    def set_state(self, checked: bool, new_value: str):
        self._traced = False
        self.new_value.set(new_value)
        self.checked.set(checked)
        self._traced = True

    def grid(self, row: int):
        self.cb.grid(row=row, column=0, padx=(8, 2), pady=1)
        self.l_coord.grid(row=row, column=1, sticky="w", pady=1)
        self.l_label.grid(row=row, column=2, sticky="w", padx=4, pady=1)
        self.l_value.grid(row=row, column=3, sticky="w", padx=4, pady=1)
        self.entry.grid(row=row, column=4, sticky="ew", padx=(4, 8), pady=1)

    def grid_forget(self):
        for w in self.widgets:
            w.grid_forget()


class SpeziSmithApp:
    def __init__(self):
        self.cfg = load_config()
        self.lang = self.cfg.get("language", LANG_DE)
        if self.lang not in LANG_LABELS:
            self.lang = LANG_DE

        self.dnd_ok = False
        if DND_AVAILABLE:
            try:
                self.root = TkinterDnD.Tk()
                self.dnd_ok = True
            except Exception:
                # tkdnd nicht ladbar (fehlende Systembibliothek o. ä.)
                self.root = tk.Tk()
        else:
            self.root = tk.Tk()
        self.root.title(f"{APP_NAME} {__version__} – Spezifikationen anpassen")
        self.root.geometry("1120x780")
        self.root.minsize(980, 640)

        try:
            self.root.iconphoto(True, tk.PhotoImage(
                file=resource_path(os.path.join("assets", "icon.png"))))
        except Exception:
            pass

        sv_ttk.set_theme(self.cfg.get("theme", "light"))

        self.fields: list = []
        self.rows: list[FieldRow] = []
        self.only_active = tk.BooleanVar(value=False)
        self.filter_var = tk.StringVar()
        self.out_mode = tk.StringVar(value=self.cfg.get("out_mode", "folder"))
        self.out_dir = tk.StringVar(value=self.cfg.get("out_dir", ""))
        self.suffix = tk.StringVar(value=self.cfg.get("suffix", "_angepasst"))
        self.verify = tk.BooleanVar(value=self.cfg.get("verify", False))
        self.replace_mode = tk.StringVar(value=self.cfg.get("replace_mode", "set"))
        self.status_var = tk.StringVar(value="Bereit.")
        self.progress_var = tk.DoubleVar(value=0)
        self.counter_var = tk.StringVar(value="0 von 0 Feldern aktiv")
        self.lang_var = tk.StringVar(value=LANG_LABELS[self.lang])

        self._queue: queue.Queue = queue.Queue()
        self._worker: threading.Thread | None = None
        self._last_out_dir: str = ""

        self._build_ui()
        self._load_template(preserve={})
        self._poll_queue()

    # ------------------------------------------------------------------
    # Aufbau
    # ------------------------------------------------------------------

    def _build_ui(self):
        root = self.root
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        self._build_header(root)

        main = ttk.Frame(root, padding=10)
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=3)
        main.rowconfigure(1, weight=0)
        main.rowconfigure(2, weight=1)

        self._build_step1(main)
        self._build_step2(main)
        self._build_step3(main)
        self._build_statusbar(root)

    def _build_header(self, root):
        header = ttk.Frame(root, padding=(12, 10, 12, 4))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        try:
            img = tk.PhotoImage(file=resource_path(LOGO_PNG))
            img = img.subsample(max(1, img.height() // 56))
            self._logo_img = img  # Referenz halten!
            card = tk.Label(header, image=img, background="white",
                            relief="groove", borderwidth=1)
            card.grid(row=0, column=0, rowspan=2, padx=(0, 14), sticky="w")
        except Exception:
            pass

        ttk.Label(header, text=APP_NAME,
                  font=("Segoe UI", 20, "bold")).grid(row=0, column=1, sticky="w")
        ttk.Label(header,
                  text="Felder wählen · neue Werte eintragen · Spezifikationen "
                       "hereinziehen – fertig.",
                  foreground="#666666").grid(row=1, column=1, sticky="w")

        ttk.Label(header, text="Sprache der Vorlage:").grid(row=0, column=2,
                                                            sticky="e", padx=(10, 4))
        lang_box = ttk.Combobox(header, textvariable=self.lang_var, width=10,
                                state="readonly",
                                values=list(LANG_LABELS.values()))
        lang_box.grid(row=0, column=3, sticky="e")
        lang_box.bind("<<ComboboxSelected>>", self._on_language_change)

        self.theme_btn = ttk.Button(header, text="🌙 Dunkel", width=11,
                                    command=self._toggle_theme)
        if sv_ttk.get_theme() == "dark":
            self.theme_btn.configure(text="☀ Hell")
        self.theme_btn.grid(row=0, column=4, sticky="e", padx=(10, 0))

    def _build_step1(self, main):
        frame = ttk.LabelFrame(
            main, text=" 1 · Felder auswählen und neue Werte eintragen ", padding=8)
        frame.grid(row=0, column=0, sticky="nsew", pady=(4, 6))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        # Werkzeugleiste
        bar = ttk.Frame(frame)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(bar, text="🔎").pack(side="left")
        filter_entry = ttk.Entry(bar, textvariable=self.filter_var, width=30)
        filter_entry.pack(side="left", padx=(2, 10))
        ToolTip(filter_entry, "Felder nach Zelle, Beschriftung oder Wert filtern")
        self.filter_var.trace_add("write", lambda *_: self._apply_filter())

        ttk.Button(bar, text="Alle", width=7,
                   command=lambda: self._set_all_visible(True)).pack(side="left")
        ttk.Button(bar, text="Keine", width=7,
                   command=lambda: self._set_all_visible(False)).pack(side="left", padx=4)
        ttk.Button(bar, text="Eingaben leeren",
                   command=self._clear_inputs).pack(side="left", padx=(0, 10))
        ttk.Checkbutton(bar, text="Nur aktive anzeigen", variable=self.only_active,
                        command=self._apply_filter).pack(side="left")
        ttk.Label(bar, textvariable=self.counter_var,
                  foreground=ORANGE, font=("Segoe UI", 10, "bold")
                  ).pack(side="right")

        # Scrollbare Feldliste
        list_frame = ttk.Frame(frame)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.canvas = tk.Canvas(list_frame, highlightthickness=0)
        self._style_canvases()
        scroll = ttk.Scrollbar(list_frame, orient="vertical",
                               command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.rows_frame = ttk.Frame(self.canvas)
        self.rows_frame.columnconfigure(2, weight=1)
        self.rows_frame.columnconfigure(3, weight=1)
        self.rows_frame.columnconfigure(4, weight=0)
        self._canvas_window = self.canvas.create_window(
            (0, 0), window=self.rows_frame, anchor="nw")
        self.rows_frame.bind("<Configure>", self._on_rows_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)       # Windows
        self.canvas.bind_all("<Button-4>", self._on_mousewheel_linux)   # Linux
        self.canvas.bind_all("<Button-5>", self._on_mousewheel_linux)

        # Spaltenköpfe
        heads = ("", "Zelle", "Feld / Beschriftung", "Aktueller Wert", "Neuer Wert")
        for col, text in enumerate(heads):
            ttk.Label(self.rows_frame, text=text,
                      font=("Segoe UI", 9, "bold")).grid(
                row=0, column=col, sticky="w",
                padx=(8 if col == 0 else 4, 4), pady=(0, 4))

    def _style_canvases(self):
        """Canvas-Hintergruende passend zum Theme setzen."""
        dark = sv_ttk.get_theme() == "dark"
        bg = "#1c1c1e" if dark else "#ffffff"
        self._canvas_bg = bg
        self._canvas_fg = "#F2F2F5" if dark else NAVY
        self._canvas_muted = "#9aa0ac" if dark else "#666666"
        for c in (getattr(self, "canvas", None), getattr(self, "drop_canvas", None)):
            if c is not None:
                c.configure(background=bg)

    def _build_step2(self, main):
        frame = ttk.LabelFrame(main, text=" 2 · Ausgabe wählen ", padding=8)
        frame.grid(row=1, column=0, sticky="ew", pady=6)

        ttk.Radiobutton(frame, text="Zielordner:", variable=self.out_mode,
                        value="folder").grid(row=0, column=0, sticky="w")
        dir_entry = ttk.Entry(frame, textvariable=self.out_dir, width=62)
        dir_entry.grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(frame, text="Durchsuchen…",
                   command=self._browse_dir).grid(row=0, column=2, padx=(0, 8))

        ttk.Radiobutton(frame, text="Nach jeder Datei fragen („Speichern unter“)",
                        variable=self.out_mode,
                        value="ask").grid(row=1, column=0, columnspan=2,
                                          sticky="w", pady=(6, 0))

        ttk.Label(frame, text="Dateinamen-Zusatz:").grid(row=0, column=3,
                                                         sticky="e", padx=(16, 4))
        suffix_entry = ttk.Entry(frame, textvariable=self.suffix, width=14)
        suffix_entry.grid(row=0, column=4, sticky="w")
        ToolTip(suffix_entry, "Wird an den Dateinamen angehängt, z. B. "
                              "„Spezi_angepasst.xlsx“.\nLeer lassen = Originalname.")

        opt = ttk.Frame(frame)
        opt.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(8, 0))
        ttk.Label(opt, text="Ersetzungsmodus:").pack(side="left")
        rb1 = ttk.Radiobutton(opt, text="Ganze Zelle ersetzen",
                              variable=self.replace_mode, value="set")
        rb1.pack(side="left", padx=(6, 10))
        rb2 = ttk.Radiobutton(opt, text="Nur Teiltext ersetzen",
                              variable=self.replace_mode, value="replace")
        rb2.pack(side="left")
        ToolTip(rb2, "Ersetzt den aktuellen Feldtext innerhalb längerer "
                     "Zelltexte (z. B. Datum im Kopftext).")
        verify_cb = ttk.Checkbutton(
            opt, text="Nur ersetzen, wenn der aktuelle Wert dem Vorlagenwert entspricht",
            variable=self.verify)
        verify_cb.pack(side="left", padx=(20, 0))
        ToolTip(verify_cb, "Sicherheitsoption: Zellen, deren Inhalt in der "
                           "abgelegten Spezifikation vom Vorlagenwert abweicht, "
                           "werden übersprungen und im Log gemeldet.")
        frame.columnconfigure(1, weight=1)

    def _build_step3(self, main):
        frame = ttk.LabelFrame(
            main, text=" 3 · Spezifikationen verarbeiten ", padding=8)
        frame.grid(row=2, column=0, sticky="nsew", pady=(6, 4))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.drop_canvas = tk.Canvas(frame, height=92, highlightthickness=0)
        self._style_canvases()
        self.drop_canvas.grid(row=0, column=0, sticky="ew", padx=2, pady=(0, 6))
        self._draw_drop_zone(False)
        self.drop_canvas.bind("<Configure>", lambda _e: self._draw_drop_zone(False))

        if self.dnd_ok:
            self.drop_canvas.drop_target_register(DND_FILES)
            self.drop_canvas.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_canvas.dnd_bind("<<DragEnter>>",
                                      lambda _e: self._draw_drop_zone(True))
            self.drop_canvas.dnd_bind("<<DragLeave>>",
                                      lambda _e: self._draw_drop_zone(False))

        pick = ttk.Button(frame, text="📂  Dateien auswählen… (Mehrfachauswahl möglich)",
                          command=self._pick_files, style="Accent.TButton")
        pick.grid(row=1, column=0, pady=(0, 6))

        # Log
        log_frame = ttk.Frame(frame)
        log_frame.grid(row=2, column=0, sticky="nsew")
        frame.rowconfigure(2, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        cols = ("file", "result", "target")
        self.log = ttk.Treeview(log_frame, columns=cols, show="headings", height=6)
        self.log.heading("file", text="Datei")
        self.log.heading("result", text="Ergebnis")
        self.log.heading("target", text="Ziel")
        self.log.column("file", width=260, anchor="w")
        self.log.column("result", width=320, anchor="w")
        self.log.column("target", width=380, anchor="w")
        self.log.tag_configure("ok", foreground="#1E7E34")
        self.log.tag_configure("warn", foreground="#B26A00")
        self.log.tag_configure("err", foreground="#C0392B")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log.bind("<Double-1>", self._open_log_target)

        log_bar = ttk.Frame(frame)
        log_bar.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        self.open_btn = ttk.Button(log_bar, text="Zielordner öffnen",
                                   command=self._open_out_dir, state="disabled")
        self.open_btn.pack(side="left")
        ttk.Button(log_bar, text="Log leeren",
                   command=lambda: self.log.delete(*self.log.get_children())
                   ).pack(side="left", padx=6)

    def _build_statusbar(self, root):
        bar = ttk.Frame(root, padding=(10, 4))
        bar.grid(row=2, column=0, sticky="ew")
        bar.columnconfigure(1, weight=1)
        self.progress = ttk.Progressbar(bar, variable=self.progress_var,
                                        maximum=100, length=220)
        self.progress.grid(row=0, column=0, sticky="w")
        ttk.Label(bar, textvariable=self.status_var).grid(row=0, column=1,
                                                          sticky="w", padx=10)

    # ------------------------------------------------------------------
    # Feldliste
    # ------------------------------------------------------------------

    def _template_path(self) -> str:
        rel = TEMPLATE_DE if self.lang == LANG_DE else TEMPLATE_EN
        return resource_path(rel)

    def _load_template(self, preserve: dict):
        try:
            self.fields = analyze_template(self._template_path())
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Vorlage konnte nicht gelesen werden:\n{exc}")
            self.fields = []
        for row in self.rows:
            row.grid_forget()
        self.rows = []
        for i, field in enumerate(self.fields, start=1):
            row = FieldRow(self, self.rows_frame, field)
            row.grid(i)
            if field.coord in preserve:
                checked, new_value = preserve[field.coord]
                row.set_state(checked, new_value)
            self.rows.append(row)
        self._apply_filter()
        self.refresh_counter()
        self.status_var.set(
            f"Vorlage ({LANG_LABELS[self.lang]}): {len(self.fields)} Felder gefunden.")

    def _snapshot_state(self) -> dict:
        return {
            r.field.coord: (r.checked.get(), r.new_value.get())
            for r in self.rows
        }

    def _on_language_change(self, _event=None):
        label = self.lang_var.get()
        self.lang = LANG_DE if label == LANG_LABELS[LANG_DE] else LANG_EN
        self._load_template(preserve=self._snapshot_state())
        self._save_cfg()

    def _apply_filter(self):
        text = self.filter_var.get().strip().lower()
        only_active = self.only_active.get()
        visible_row = 1
        for row in self.rows:
            hay = f"{row.field.coord} {row.field.label} {row.field.value}".lower()
            show = (not text or text in hay) and \
                   (not only_active or row.checked.get())
            if show:
                row.grid(visible_row)
                visible_row += 1
            else:
                row.grid_forget()

    def _set_all_visible(self, state: bool):
        for row in self.rows:
            if row.l_coord.winfo_manager():  # nur sichtbare Zeilen
                row.checked.set(state)
        self.refresh_counter()

    def _clear_inputs(self):
        for row in self.rows:
            row.set_state(False, "")
        self.refresh_counter()

    def refresh_counter(self):
        active = sum(1 for r in self.rows if r.checked.get())
        self.counter_var.set(f"{active} von {len(self.rows)} Feldern aktiv")

    def _collect_rules(self) -> list[Rule]:
        rules = []
        mode = self.replace_mode.get()
        for row in self.rows:
            if row.checked.get():
                rules.append(Rule(
                    coord=row.field.coord,
                    new=row.new_value.get(),
                    mode=mode,
                    find=row.field.value,
                    expected=row.field.value,
                    label=row.field.label,
                ))
        return rules

    # ------------------------------------------------------------------
    # Scrollen / Drop-Zone
    # ------------------------------------------------------------------

    def _on_rows_configure(self, _event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        self.canvas.itemconfigure(self._canvas_window, width=event.width)

    def _on_mousewheel(self, event):
        if self._pointer_in_canvas():
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_mousewheel_linux(self, event):
        if self._pointer_in_canvas():
            self.canvas.yview_scroll(-1 if event.num == 4 else 1, "units")

    def _pointer_in_canvas(self) -> bool:
        widget = self.root.winfo_containing(*self.root.winfo_pointerxy())
        while widget is not None:
            if widget == self.canvas:
                return True
            widget = widget.master
        return False

    def _draw_drop_zone(self, active: bool):
        c = self.drop_canvas
        c.delete("all")
        w = max(c.winfo_width(), 400)
        h = 92
        dark = sv_ttk.get_theme() == "dark"
        color = ORANGE if active else ("#5b6270" if dark else "#8A93A6")
        fill = "#3a2c18" if (active and dark) else ("#FFF3E2" if active else self._canvas_bg)
        c.create_rectangle(4, 4, w - 4, h - 4, outline=color, width=2,
                           dash=(8, 5), fill=fill)
        line1 = "⬇  Spezifikationen hierher ziehen (.xlsx) – Mehrfachauswahl möglich"
        if not self.dnd_ok:
            line1 = "Drag & Drop nicht verfügbar – bitte die Schaltfläche unten nutzen"
        c.create_text(w / 2, h / 2 - 12, text=line1,
                      font=("Segoe UI", 11, "bold"), fill=self._canvas_fg)
        c.create_text(w / 2, h / 2 + 16,
                      text="z. B. 5 Dateien gleichzeitig aus dem Explorer hereinziehen",
                      font=("Segoe UI", 9), fill=self._canvas_muted)

    # ------------------------------------------------------------------
    # Dateiverarbeitung
    # ------------------------------------------------------------------

    def _on_drop(self, event):
        self._draw_drop_zone(False)
        paths = list(self.root.tk.splitlist(event.data))
        self._handle_files(paths)

    def _pick_files(self):
        paths = filedialog.askopenfilenames(
            parent=self.root, title="Spezifikationen auswählen (Mehrfachauswahl)",
            filetypes=FILE_TYPES)
        if paths:
            self._handle_files(list(paths))

    def _handle_files(self, paths: list[str]):
        files = [p for p in paths if p.lower().endswith((".xlsx", ".xlsm"))]
        for other in set(paths) - set(files):
            self._log(os.path.basename(other), "warn",
                      "Übersprungen – kein .xlsx/.xlsm", "")
        if not files:
            if not paths:
                return
            self.status_var.set("Keine Excel-Dateien dabei (.xlsx/.xlsm).")
            return

        rules = self._collect_rules()
        if not rules:
            messagebox.showwarning(
                APP_NAME,
                "Bitte zuerst in Schritt 1 mindestens ein Feld aktivieren\n"
                "und einen neuen Wert eintragen.")
            self.status_var.set("Keine Felder ausgewählt – nichts zu tun.")
            return

        if self.out_mode.get() == "folder":
            self._process_to_folder(files, rules)
        else:
            self._process_with_dialogs(files, rules)

    def _process_to_folder(self, files: list[str], rules: list[Rule]):
        out_dir = self.out_dir.get().strip()
        if not out_dir:
            messagebox.showinfo(APP_NAME, "Bitte zuerst in Schritt 2 einen "
                                          "Zielordner angeben.")
            self._browse_dir()
            out_dir = self.out_dir.get().strip()
            if not out_dir:
                return
        os.makedirs(out_dir, exist_ok=True)
        self._last_out_dir = out_dir
        self._save_cfg()

        if self._worker and self._worker.is_alive():
            messagebox.showinfo(APP_NAME, "Es läuft bereits eine Verarbeitung.")
            return

        self.progress_var.set(0)
        self.progress.configure(maximum=len(files))
        self.open_btn.configure(state="disabled")
        self.status_var.set(f"Verarbeite {len(files)} Datei(en)…")

        self._worker = threading.Thread(
            target=self._batch_worker,
            args=(files, rules, out_dir, self.suffix.get(), self.verify.get()),
            daemon=True)
        self._worker.start()

    def _batch_worker(self, files, rules, out_dir, suffix, verify):
        errors = []
        for i, src in enumerate(files, start=1):
            lang = detect_language(src)
            if lang and lang != self.lang:
                self._queue.put(("log", os.path.basename(src), "warn",
                                 f"Sprache passt nicht zur Auswahl "
                                 f"({LANG_LABELS.get(lang, lang)}) – übersprungen", ""))
            else:
                target = unique_target_path(out_dir, os.path.basename(src), suffix)
                res = apply_to_file(src, target, rules, verify=verify)
                if res.ok:
                    msg = f"{res.changed_count} Feld(er) geändert"
                    if res.skipped_count:
                        msg += f", {res.skipped_count} übersprungen"
                    self._queue.put(("log", os.path.basename(src), "ok", msg, target))
                else:
                    errors.append(f"{os.path.basename(src)}: {res.error}")
                    self._queue.put(("log", os.path.basename(src), "err",
                                     f"Fehler: {res.error}", ""))
            self._queue.put(("progress", i))
        self._queue.put(("done", len(files), errors))

    def _process_with_dialogs(self, files: list[str], rules: list[Rule]):
        errors = []
        for i, src in enumerate(files, start=1):
            name = os.path.basename(src)
            lang = detect_language(src)
            if lang and lang != self.lang:
                self._log(name, "warn",
                          f"Sprache passt nicht zur Auswahl "
                          f"({LANG_LABELS.get(lang, lang)}) – übersprungen", "")
                continue
            stem, ext = os.path.splitext(name)
            target = filedialog.asksaveasfilename(
                parent=self.root, title=f"Geänderte Spezifikation speichern: {name}",
                initialfile=f"{stem}{self.suffix.get()}{ext}",
                defaultextension=ext, filetypes=FILE_TYPES)
            if not target:
                self._log(name, "warn", "Speichern abgebrochen", "")
                continue
            res = apply_to_file(src, target, rules, verify=self.verify.get())
            if res.ok:
                self._log(name, "ok", f"{res.changed_count} Feld(er) geändert", target)
                self._last_out_dir = os.path.dirname(target)
            else:
                errors.append(f"{name}: {res.error}")
                self._log(name, "err", f"Fehler: {res.error}", "")
            self.progress_var.set(i)
        self.status_var.set(f"Fertig – {len(files)} Datei(en) verarbeitet.")
        if errors:
            messagebox.showerror(APP_NAME, "Fehler bei:\n\n" + "\n".join(errors))
        self._save_cfg()

    def _poll_queue(self):
        try:
            while True:
                item = self._queue.get_nowait()
                kind = item[0]
                if kind == "log":
                    _, name, tag, msg, target = item
                    self._log(name, tag, msg, target)
                elif kind == "progress":
                    self.progress_var.set(item[1])
                    self.status_var.set(f"Verarbeite… {item[1]} Datei(en) fertig.")
                elif kind == "done":
                    _, total, errors = item
                    self.progress_var.set(total)
                    self.status_var.set(f"Fertig – {total} Datei(en) verarbeitet.")
                    self.open_btn.configure(state="normal")
                    if errors:
                        messagebox.showerror(
                            APP_NAME, "Fehler bei:\n\n" + "\n".join(errors))
        except queue.Empty:
            pass
        self.root.after(60, self._poll_queue)

    # ------------------------------------------------------------------
    # Diverses
    # ------------------------------------------------------------------

    def _log(self, name: str, tag: str, msg: str, target: str):
        self.log.insert("", "end", values=(name, msg, target), tags=(tag,))
        self.log.see(self.log.get_children()[-1])

    def _browse_dir(self):
        path = filedialog.askdirectory(parent=self.root, title="Zielordner wählen")
        if path:
            self.out_dir.set(path)
            self.out_mode.set("folder")

    def _open_out_dir(self):
        if not self._last_out_dir:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self._last_out_dir)  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self._last_out_dir])
            else:
                subprocess.Popen(["xdg-open", self._last_out_dir])
        except Exception as exc:
            messagebox.showerror(APP_NAME, f"Ordner konnte nicht geöffnet werden:\n{exc}")

    def _open_log_target(self, _event):
        sel = self.log.selection()
        if not sel:
            return
        target = self.log.item(sel[0], "values")[2]
        if target and os.path.exists(os.path.dirname(target)):
            self._last_out_dir = os.path.dirname(target)
            self._open_out_dir()

    def _toggle_theme(self):
        new = "dark" if sv_ttk.get_theme() == "light" else "light"
        sv_ttk.set_theme(new)
        self.theme_btn.configure(text="☀ Hell" if new == "dark" else "🌙 Dunkel")
        self._style_canvases()
        self._draw_drop_zone(False)
        self._save_cfg()

    def _save_cfg(self):
        save_config({
            "language": self.lang,
            "theme": sv_ttk.get_theme(),
            "out_mode": self.out_mode.get(),
            "out_dir": self.out_dir.get(),
            "suffix": self.suffix.get(),
            "verify": self.verify.get(),
            "replace_mode": self.replace_mode.get(),
        })

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        self._save_cfg()
        self.root.destroy()


def run():
    SpeziSmithApp().run()
