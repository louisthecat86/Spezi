"""
SpeziSmith - Grafische Oberfläche

Modernes UI (Sun-Valley-Theme) mit:
- Sprachauswahl Deutsch/English (steuert die analysierte Vorlage)
- maßstabsgetreuer, klickbarer Vorlagen-Ansicht (WYSIWYG): Feld anklicken,
  Änderung eintragen, Ergebnis direkt im Layout sehen
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
    HF_PREFIX,
    Rule,
    apply_to_file,
    detect_language,
    layout_fingerprints,
    layout_similarity,
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
from .sheetview import SheetView, hf_friendly_name

NAVY = "#1F2942"
ORANGE = "#F08C21"

LANG_DE = "de"
LANG_EN = "en"
LANG_LABELS = {LANG_DE: "Deutsch", LANG_EN: "English"}

FILE_TYPES = [("Excel-Spezifikation", "*.xlsx *.xlsm")]


class ToolTip:
    """Kleiner Hover-Tooltip für Widgets."""

    def __init__(self, widget: tk.Widget, text: str):
        self.text = text
        self.tip: tk.Toplevel | None = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _event=None):
        if self.tip is not None or not self.text:
            return
        self.tip = tk.Toplevel()
        self.tip.overrideredirect(True)
        self.tip.attributes("-topmost", True)
        label = tk.Label(
            self.tip, text=self.text, background="#FFFFE0", relief="solid",
            borderwidth=1, padx=6, pady=3, wraplength=340, justify="left",
        )
        label.pack()
        x = self.tip.winfo_pointerx() + 12
        y = self.tip.winfo_pointery() + 16
        self.tip.geometry(f"+{x}+{y}")

    def _hide(self, _event=None):
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


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
        self.root.geometry("1240x800")
        self.root.minsize(1020, 660)

        try:
            self.root.iconphoto(True, tk.PhotoImage(
                file=resource_path(os.path.join("assets", "icon.png"))))
        except Exception:
            pass

        sv_ttk.set_theme(self.cfg.get("theme", "light"))

        self.search_var = tk.StringVar()
        self.out_mode = tk.StringVar(value=self.cfg.get("out_mode", "folder"))
        self.out_dir = tk.StringVar(value=self.cfg.get("out_dir", ""))
        self.suffix = tk.StringVar(value=self.cfg.get("suffix", "_angepasst"))
        self.verify = tk.BooleanVar(value=self.cfg.get("verify", False))
        self.layout_check = tk.BooleanVar(value=self.cfg.get("layout_check", True))
        self.replace_mode = tk.StringVar(value=self.cfg.get("replace_mode", "set"))
        self.zoom_var = tk.StringVar(value=self.cfg.get("zoom", "100 %"))
        self.status_var = tk.StringVar(value="Bereit.")
        self.progress_var = tk.DoubleVar(value=0)
        self.counter_var = tk.StringVar(value="Keine Änderungen")
        self.lang_var = tk.StringVar(value=LANG_LABELS[self.lang])
        self._layout_ref: dict = {}

        # Bearbeitungsleiste (Schritt 1)
        self.sel_coord_var = tk.StringVar(value="–")
        self.sel_orig_var = tk.StringVar(value="")
        self.sel_new_var = tk.StringVar()
        self.sel_clear_var = tk.BooleanVar(value=False)
        self._edit_guard = False
        self._selected_coord: str | None = None

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

        main = ttk.Frame(root, padding=(10, 4, 10, 6))
        main.grid(row=1, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=3)
        main.rowconfigure(2, weight=1)

        self._build_step1(main)
        self._build_step2(main)
        self._build_step3(main)
        self._build_statusbar(root)

    def _build_header(self, root):
        header = ttk.Frame(root, padding=(12, 8, 12, 2))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)

        try:
            img = tk.PhotoImage(file=resource_path(LOGO_PNG))
            img = img.subsample(max(1, img.height() // 52))
            self._logo_img = img  # Referenz halten!
            card = tk.Label(header, image=img, background="white",
                            relief="groove", borderwidth=1)
            card.grid(row=0, column=0, rowspan=2, padx=(0, 14), sticky="w")
        except Exception:
            pass

        ttk.Label(header, text=APP_NAME,
                  font=("Segoe UI", 20, "bold")).grid(row=0, column=1, sticky="w")
        ttk.Label(header,
                  text="Feld in der Vorlage anklicken · Änderung eintragen · "
                       "Spezifikationen hereinziehen – fertig.",
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

    # ------------------------------------------------------------------
    # Schritt 1: Vorlagen-Ansicht + Bearbeitungsleiste
    # ------------------------------------------------------------------

    def _build_step1(self, main):
        frame = ttk.LabelFrame(
            main, text=" 1 · Feld anklicken und Änderung eintragen ", padding=6)
        frame.grid(row=0, column=0, sticky="nsew", pady=(2, 6))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        # Werkzeugleiste: Suche + Verwerfen + Zähler
        bar = ttk.Frame(frame)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Label(bar, text="🔎").pack(side="left")
        search = ttk.Entry(bar, textvariable=self.search_var, width=30)
        search.pack(side="left", padx=(2, 4))
        search.bind("<Return>", lambda _e: self._search_next())
        ToolTip(search, "Zelle (z. B. „A5“) oder Text suchen – Enter springt "
                        "zum nächsten Treffer")
        ttk.Button(bar, text="Weitersuchen", width=12,
                   command=self._search_next).pack(side="left")
        self.search_info_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.search_info_var,
                  foreground="#666666").pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Alle Änderungen verwerfen",
                   command=self._clear_all_changes).pack(side="left", padx=(16, 0))

        ttk.Label(bar, text="Zoom:").pack(side="left", padx=(16, 2))
        zoom_box = ttk.Combobox(bar, textvariable=self.zoom_var, width=6,
                                state="readonly",
                                values=["60 %", "70 %", "80 %", "90 %", "100 %",
                                        "110 %", "125 %", "150 %"])
        zoom_box.pack(side="left")
        zoom_box.bind("<<ComboboxSelected>>", self._on_zoom_change)
        ToolTip(zoom_box, "Ansicht verkleinern, um die ganze Seite zu sehen")

        ttk.Label(bar, textvariable=self.counter_var,
                  foreground=ORANGE, font=("Segoe UI", 10, "bold")
                  ).pack(side="right")

        # Klickbare Vorlagen-Ansicht
        self.sheet = SheetView(frame, on_select=self._on_cell_select)
        self.sheet.grid(row=1, column=0, sticky="nsew")

        # Bearbeitungsleiste
        edit = ttk.Frame(frame, padding=(2, 6, 2, 0))
        edit.grid(row=2, column=0, sticky="ew")
        edit.columnconfigure(3, weight=1)

        ttk.Label(edit, text="Zelle:").grid(row=0, column=0, sticky="w")
        ttk.Label(edit, textvariable=self.sel_coord_var, width=7,
                  font=("Consolas", 11, "bold"), foreground=NAVY
                  ).grid(row=0, column=1, sticky="w")
        self.orig_label = ttk.Label(edit, textvariable=self.sel_orig_var,
                                    foreground="#666666", width=44)
        self.orig_label.grid(row=0, column=2, sticky="w", padx=(4, 10))

        ttk.Label(edit, text="Neuer Wert:").grid(row=0, column=3, sticky="e")
        self.new_entry = ttk.Entry(edit, textvariable=self.sel_new_var, width=42)
        self.new_entry.grid(row=0, column=4, sticky="ew", padx=(4, 6))
        self.sel_new_var.trace_add("write", self._on_new_value_typed)

        self.clear_cb = ttk.Checkbutton(edit, text="Feld leeren",
                                        variable=self.sel_clear_var,
                                        command=self._on_clear_toggled)
        self.clear_cb.grid(row=0, column=5)
        ToolTip(self.clear_cb, "Den Inhalt des gewählten Feldes in allen "
                               "abgelegten Spezifikationen entfernen")
        ttk.Button(edit, text="Zurücknehmen", width=13,
                   command=self._revert_current).grid(row=0, column=6, padx=(6, 0))

        self._edit_widgets = (self.new_entry, self.clear_cb)
        self._set_edit_enabled(False)

    def _set_edit_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for w in self._edit_widgets:
            w.configure(state=state)

    def _on_cell_select(self, coord):
        if coord is None:
            self._set_edit_enabled(False)
            self._selected_coord = None
            self.sel_coord_var.set("–")
            self.sel_orig_var.set("")
            return
        self._selected_coord = coord
        cell = self.sheet.cells[coord]
        orig = cell.orig_text
        is_hf = coord.startswith(HF_PREFIX)
        self.sel_coord_var.set(hf_friendly_name(coord) if is_hf else coord)
        short = orig if len(orig) <= 60 else orig[:60] + "…"
        self.sel_orig_var.set(f"aktuell: {short or '(leer)'}")
        tip = orig or "(leer)"
        if is_hf:
            tip += "\n\nHinweis: &P = Seitenzahl, &N = Seiten gesamt, &G = Grafik"
        ToolTip(self.orig_label, tip)
        self._edit_guard = True
        self.sel_new_var.set(cell.new_text or "")
        self.sel_clear_var.set(cell.new_text == "")
        self._edit_guard = False
        self._set_edit_enabled(True)
        if cell.new_text == "":
            self.new_entry.configure(state="disabled")
        self.new_entry.focus_set()

    def _current_coord(self):
        coord = self._selected_coord
        return coord if coord in self.sheet.cells else None

    def _on_new_value_typed(self, *_):
        if self._edit_guard:
            return
        coord = self._current_coord()
        if not coord or self.sel_clear_var.get():
            return
        text = self.sel_new_var.get()
        self.sheet.set_new_value(coord, text if text else None)
        self._refresh_counter()

    def _on_clear_toggled(self):
        coord = self._current_coord()
        if not coord:
            return
        if self.sel_clear_var.get():
            self._edit_guard = True
            self.sel_new_var.set("")
            self._edit_guard = False
            self.new_entry.configure(state="disabled")
            self.sheet.set_new_value(coord, "")
        else:
            self.new_entry.configure(state="normal")
            text = self.sel_new_var.get()
            self.sheet.set_new_value(coord, text if text else None)
        self._refresh_counter()

    def _revert_current(self):
        coord = self._current_coord()
        if not coord:
            return
        self._edit_guard = True
        self.sel_new_var.set("")
        self.sel_clear_var.set(False)
        self._edit_guard = False
        self.new_entry.configure(state="normal")
        self.sheet.set_new_value(coord, None)
        self._refresh_counter()

    def _clear_all_changes(self):
        self.sheet.clear_all()
        self._edit_guard = True
        self.sel_new_var.set("")
        self.sel_clear_var.set(False)
        self._edit_guard = False
        self.new_entry.configure(state="normal")
        self._refresh_counter()

    def _refresh_counter(self):
        count = len(self.sheet.changed())
        self.counter_var.set(
            f"{count} Feld(er) geändert" if count else "Keine Änderungen")

    def _search_next(self):
        text = self.search_var.get()
        if not text.strip():
            self.sheet.search("")
            self.search_info_var.set("")
            return
        if self.sheet._match_idx < 0:
            total = self.sheet.search(text)
            if total == 0:
                self.search_info_var.set("nichts gefunden")
                return
        coord = self.sheet.next_match()
        if coord:
            self.search_info_var.set(
                f"Treffer {self.sheet._match_idx + 1}/{len(self.sheet._matches)}"
                f" – {coord}")

    # ------------------------------------------------------------------
    # Schritt 2: Ausgabe
    # ------------------------------------------------------------------

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
                           "werden übersprungen und im Log gemeldet. Leere "
                           "Eingabefelder werden nur befüllt, wenn sie in der "
                           "Zieldatei leer sind.")
        layout_cb = ttk.Checkbutton(
            opt, text="Layout-Prüfung (empfohlen)", variable=self.layout_check)
        layout_cb.pack(side="left", padx=(20, 0))
        ToolTip(layout_cb, "Prüft jede abgelegte Datei vor dem Ersetzen: Weicht "
                           "ihr Layout deutlich von der Vorlage ab (z. B. andere "
                           "Zeilenanzahl), wird sie übersprungen statt falsch "
                           "befüllt - mit Warnung im Log.")
        frame.columnconfigure(1, weight=1)

    # ------------------------------------------------------------------
    # Schritt 3: Verarbeiten
    # ------------------------------------------------------------------

    def _build_step3(self, main):
        frame = ttk.LabelFrame(
            main, text=" 3 · Spezifikationen verarbeiten ", padding=8)
        frame.grid(row=2, column=0, sticky="nsew", pady=(6, 2))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(2, weight=1)

        self.drop_canvas = tk.Canvas(frame, height=76, highlightthickness=0)
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

        log_frame = ttk.Frame(frame)
        log_frame.grid(row=2, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        cols = ("file", "result", "target")
        self.log = ttk.Treeview(log_frame, columns=cols, show="headings", height=5)
        self.log.heading("file", text="Datei")
        self.log.heading("result", text="Ergebnis")
        self.log.heading("target", text="Ziel")
        self.log.column("file", width=260, anchor="w")
        self.log.column("result", width=320, anchor="w")
        self.log.column("target", width=420, anchor="w")
        self.log.tag_configure("ok", foreground="#1E7E34")
        self.log.tag_configure("warn", foreground="#B26A00")
        self.log.tag_configure("err", foreground="#C0392B")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log.bind("<Double-1>", self._open_log_target)
        ToolTip(self.log, "Doppelklick auf einen Eintrag öffnet den Zielordner")

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

    def _style_canvases(self):
        """Canvas-Hintergruende passend zum Theme setzen."""
        dark = sv_ttk.get_theme() == "dark"
        self._canvas_bg = "#1c1c1e" if dark else "#ffffff"
        self._canvas_fg = "#F2F2F5" if dark else NAVY
        self._canvas_muted = "#9aa0ac" if dark else "#666666"
        if getattr(self, "drop_canvas", None) is not None:
            self.drop_canvas.configure(background=self._canvas_bg)

    def _draw_drop_zone(self, active: bool):
        c = self.drop_canvas
        c.delete("all")
        w = max(c.winfo_width(), 400)
        h = 76
        dark = sv_ttk.get_theme() == "dark"
        color = ORANGE if active else ("#5b6270" if dark else "#8A93A6")
        fill = "#3a2c18" if (active and dark) else ("#FFF3E2" if active else self._canvas_bg)
        c.create_rectangle(4, 4, w - 4, h - 4, outline=color, width=2,
                           dash=(8, 5), fill=fill)
        line1 = "⬇  Spezifikationen hierher ziehen (.xlsx) – Mehrfachauswahl möglich"
        if not self.dnd_ok:
            line1 = "Drag & Drop nicht verfügbar – bitte die Schaltfläche unten nutzen"
        c.create_text(w / 2, h / 2 - 10, text=line1,
                      font=("Segoe UI", 11, "bold"), fill=self._canvas_fg)
        c.create_text(w / 2, h / 2 + 14,
                      text="z. B. 5 Dateien gleichzeitig aus dem Explorer hereinziehen",
                      font=("Segoe UI", 9), fill=self._canvas_muted)

    # ------------------------------------------------------------------
    # Vorlage / Sprache
    # ------------------------------------------------------------------

    def _template_path(self) -> str:
        rel = TEMPLATE_DE if self.lang == LANG_DE else TEMPLATE_EN
        return resource_path(rel)

    def _zoom_scale(self) -> float:
        try:
            return max(40, min(200, int(self.zoom_var.get().rstrip(" %")))) / 100
        except ValueError:
            return 1.0

    def _on_zoom_change(self, _event=None):
        self._load_template(preserve=self.sheet.changed())
        self._save_cfg()

    def _load_template(self, preserve: dict):
        self.status_var.set("Analysiere Vorlage…")
        self.root.update_idletasks()
        try:
            count = self.sheet.load_template(self._template_path(), keep=preserve,
                                             scale=self._zoom_scale())
            self._layout_ref = layout_fingerprints(self._template_path())
        except Exception as exc:
            messagebox.showerror(APP_NAME,
                                 f"Vorlage konnte nicht gelesen werden:\n{exc}")
            count = 0
        self._refresh_counter()
        self.search_info_var.set("")
        self.search_var.set("")
        self.status_var.set(
            f"Vorlage ({LANG_LABELS[self.lang]}): {count} Felder – "
            f"bitte ein Feld anklicken.")

    def _on_language_change(self, _event=None):
        label = self.lang_var.get()
        self.lang = LANG_DE if label == LANG_LABELS[LANG_DE] else LANG_EN
        self._load_template(preserve=self.sheet.changed())
        self._save_cfg()

    def _collect_rules(self) -> list[Rule]:
        rules = []
        mode = self.replace_mode.get()
        for coord, new in self.sheet.changed().items():
            orig = self.sheet.orig_value(coord)
            anchor_text, anchor_coord = self.sheet.anchor_of(coord)
            if coord.startswith(HF_PREFIX):
                label = hf_friendly_name(coord)
                mode = "set"  # Kopf-/Fußzeile: format-schonender Volltext-Ersatz
            else:
                label = orig[:40] or coord
            rules.append(Rule(
                coord=coord, new=new, mode=mode,
                find=orig, expected=orig,
                label=label,
                anchor_text=anchor_text, anchor_coord=anchor_coord,
            ))
        return rules

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
                "Bitte zuerst in Schritt 1 mindestens ein Feld anklicken\n"
                "und einen neuen Wert eintragen.")
            self.status_var.set("Keine Felder geändert – nichts zu tun.")
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
            args=(files, rules, out_dir, self.suffix.get(), self.verify.get(),
                  self.layout_check.get(), self._layout_ref),
            daemon=True)
        self._worker.start()

    def _batch_worker(self, files, rules, out_dir, suffix, verify,
                      layout_check, layout_ref):
        errors = []
        for i, src in enumerate(files, start=1):
            lang = detect_language(src)
            if lang and lang != self.lang:
                self._queue.put(("log", os.path.basename(src), "warn",
                                 f"Sprache passt nicht zur Auswahl "
                                 f"({LANG_LABELS.get(lang, lang)}) – übersprungen", ""))
            elif layout_check and layout_ref:
                sim = layout_similarity(layout_ref, layout_fingerprints(src))
                if sim < 0.75:
                    self._queue.put(("log", os.path.basename(src), "warn",
                                     f"Layout weicht ab (nur {sim:.0%} Übereinstimmung"
                                     f" mit der Vorlage) – übersprungen", ""))
                else:
                    self._apply_one(src, out_dir, suffix, rules, verify, errors)
            else:
                self._apply_one(src, out_dir, suffix, rules, verify, errors)
            self._queue.put(("progress", i))
        self._queue.put(("done", len(files), errors))

    def _apply_one(self, src, out_dir, suffix, rules, verify, errors):
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

    def _process_with_dialogs(self, files: list[str], rules: list[Rule]):
        errors = []
        layout_check = self.layout_check.get()
        for i, src in enumerate(files, start=1):
            name = os.path.basename(src)
            lang = detect_language(src)
            if lang and lang != self.lang:
                self._log(name, "warn",
                          f"Sprache passt nicht zur Auswahl "
                          f"({LANG_LABELS.get(lang, lang)}) – übersprungen", "")
                continue
            if layout_check and self._layout_ref:
                sim = layout_similarity(self._layout_ref, layout_fingerprints(src))
                if sim < 0.75:
                    self._log(name, "warn",
                              f"Layout weicht ab (nur {sim:.0%} Übereinstimmung "
                              f"mit der Vorlage) – übersprungen", "")
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
            "layout_check": self.layout_check.get(),
            "replace_mode": self.replace_mode.get(),
            "zoom": self.zoom_var.get(),
        })

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.mainloop()

    def _on_close(self):
        self._save_cfg()
        self.root.destroy()


def run():
    SpeziSmithApp().run()
