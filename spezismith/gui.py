"""
SpeziSmith - Grafische Oberfläche

Modernes UI (Sun-Valley-Theme) mit:
- Vorlage hochladen (beliebige korrekte Spezifikation) statt fester Vorlage
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
import time
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
    context_above,
    split_coord,
    layout_fingerprints,
    layout_similarity,
    template_row_texts,
    LAYOUT_MIN_SIMILARITY,
    unique_target_path,
)
from .resources import (
    LOGO_PNG,
    load_config,
    resource_path,
    save_config,
)
from .sheetview import SheetView, hf_friendly_name

NAVY = "#1F2942"
ORANGE = "#F08C21"

FILE_TYPES = [("Excel-Spezifikation", "*.xlsx *.xlsm")]


def _changed_lines(orig: str, new: str):
    """Mehrzeilige Zelle: nur die geänderten Zeilen als (find, ersatz).

    So bleiben abweichende Nachbarzeilen in der Zieldatei unangetastet.
    None = nicht eindeutig teilbar (dann wird die ganze Zelle gesetzt).
    """
    a, b = orig.split("\n"), new.split("\n")
    i = 0
    while i < min(len(a), len(b)) and a[i] == b[i]:
        i += 1
    j = 0
    while j < min(len(a), len(b)) - i and a[-1 - j] == b[-1 - j]:
        j += 1
    mid_a, mid_b = a[i:len(a) - j], b[i:len(b) - j]
    if not mid_a or not mid_b:
        return None
    return "\n".join(mid_a), "\n".join(mid_b), i, len(mid_a)


class VarText(tk.Text):
    """Mehrzeiliges Eingabefeld, das wie ein Entry an eine StringVar gebunden ist."""

    def __init__(self, master, var: tk.StringVar, height: int = 3):
        super().__init__(master, height=height, width=30, wrap="word",
                         undo=True, font=("Segoe UI", 10), borderwidth=1,
                         relief="solid", padx=4, pady=2)
        self._var = var
        self.bind("<KeyRelease>", self._to_var)
        self.bind("<<Paste>>", lambda _e: self.after_idle(self._to_var))
        self.bind("<<Cut>>", lambda _e: self.after_idle(self._to_var))
        var.trace_add("write", self._from_var)

    def _to_var(self, _event=None):
        text = self.get("1.0", "end-1c")
        if text != self._var.get():
            self._var.set(text)

    def _from_var(self, *_):
        text = self._var.get()
        if text != self.get("1.0", "end-1c"):
            state = str(self.cget("state"))
            self.configure(state="normal")
            self.delete("1.0", "end")
            self.insert("1.0", text)
            self.configure(state=state)


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
        self.template_path: str | None = None
        self._row_texts: dict[int, str] = {}

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
        self.root.title(f"{APP_NAME} {__version__} – Spezifikationen an Vorlage angleichen")
        self.root.geometry("1400x860")
        self.root.minsize(1100, 680)

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
        self.layout_check = tk.BooleanVar(value=self.cfg.get("layout_check", False))
        self.replace_mode = tk.StringVar(value=self.cfg.get("replace_mode", "set"))
        self.zoom_var = tk.StringVar(value=self.cfg.get("zoom", "Ganze Seite"))
        self.status_var = tk.StringVar(value="Bereit.")
        self.progress_var = tk.DoubleVar(value=0)
        self.counter_var = tk.StringVar(value="Keine Änderungen")
        self.template_var = tk.StringVar(value="Keine Vorlage geladen")
        self.adopt_var = tk.BooleanVar(value=False)
        self._layout_ref: dict = {}
        self._progress_started_at: float | None = None
        self._progress_total = 0
        self._progress_done = 0
        self._progress_current = ""
        self._progress_visible = False
        self._progress_reveal_job: str | None = None

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
        self._show_placeholder(True)
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
        main.rowconfigure(0, weight=1)

        # Links die Seitenansicht (volle Höhe), rechts Ausgabe + Verarbeitung
        side = ttk.Frame(main)
        side.grid(row=0, column=1, sticky="ns", padx=(8, 0))
        side.columnconfigure(0, weight=1)
        side.rowconfigure(1, weight=1)

        self._build_step1(main)
        self._build_step2(side)
        self._build_step3(side)
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
                  text="Vorlage hochladen · Felder markieren · "
                       "Spezifikationen hereinziehen – alle werden auf die Vorlage angeglichen.",
                  foreground="#666666").grid(row=1, column=1, sticky="w")

        ttk.Label(header, textvariable=self.template_var,
                  foreground="#666666").grid(row=0, column=2, sticky="e",
                                             padx=(10, 6))
        ttk.Button(header, text="📂 Vorlage hochladen…", style="Accent.TButton",
                   command=self._pick_template).grid(row=0, column=3, sticky="e")

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
            main, text=" 1 · Vorlage: Felder anklicken, die in alle Spezifikationen gelten sollen ", padding=6)
        frame.grid(row=0, column=0, sticky="nsew", pady=(2, 6))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        # Werkzeugleiste: Suche + Verwerfen + Zähler
        bar = ttk.Frame(frame)
        bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
        ttk.Label(bar, text="🔎").pack(side="left")
        search = ttk.Entry(bar, textvariable=self.search_var, width=16)
        search.pack(side="left", padx=(2, 4))
        search.bind("<Return>", lambda _e: self._search_next())
        ToolTip(search, "Zelle (z. B. „A5“) oder Text suchen – Enter springt "
                        "zum nächsten Treffer")
        ttk.Button(bar, text="Weitersuchen", width=12,
                   command=self._search_next).pack(side="left")
        self.search_info_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.search_info_var,
                  foreground="#666666").pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Alle verwerfen",
                   command=self._clear_all_changes).pack(side="left", padx=(8, 0))

        ttk.Label(bar, text="Zoom:").pack(side="left", padx=(8, 2))
        zoom_box = ttk.Combobox(bar, textvariable=self.zoom_var, width=11,
                                state="readonly",
                                values=["Ganze Seite", "60 %", "70 %", "80 %", "90 %", "100 %",
                                        "110 %", "125 %", "150 %"])
        zoom_box.pack(side="left")
        zoom_box.bind("<<ComboboxSelected>>", self._on_zoom_change)
        ToolTip(zoom_box, "„Ganze Seite“ passt die Seite ohne Scrollen ein; Seitenwechsel unter der Ansicht")

        ttk.Label(bar, textvariable=self.counter_var,
                  foreground=ORANGE, font=("Segoe UI", 10, "bold")
                  ).pack(side="right")

        # Klickbare Vorlagen-Ansicht
        self.sheet = SheetView(frame, on_select=self._on_cell_select)
        self.sheet.grid(row=1, column=0, sticky="nsew")

        self.placeholder = tk.Frame(frame, highlightthickness=0)
        ttk.Label(self.placeholder,
                  text="Noch keine Vorlage geladen",
                  font=("Segoe UI", 16, "bold")).pack(pady=(0, 6))
        ttk.Label(self.placeholder,
                  text="Lade zuerst eine korrekte Spezifikation als Vorlage hoch.\n"
                       "SpeziSmith liest alle Felder samt Beschriftung ein; danach "
                       "werden alle hereingezogenen Spezifikationen\n"
                       "auf diese Vorlage angeglichen (fehlende Felder werden "
                       "eingefügt, gewählte Werte übernommen).",
                  justify="center", foreground="#666666").pack(pady=(0, 12))
        ttk.Button(self.placeholder, text="📂  Vorlage hochladen…",
                   style="Accent.TButton",
                   command=self._pick_template).pack()
        if self.dnd_ok:
            self.placeholder.drop_target_register(DND_FILES)
            self.placeholder.dnd_bind("<<Drop>>", self._on_template_drop)

        # Bearbeitungsleiste
        edit = ttk.Frame(frame, padding=(2, 6, 2, 0))
        edit.grid(row=2, column=0, sticky="ew")
        edit.columnconfigure(3, weight=1)

        ttk.Label(edit, text="Zelle:").grid(row=0, column=0, sticky="w")
        ttk.Label(edit, textvariable=self.sel_coord_var, width=7,
                  font=("Consolas", 11, "bold"), foreground=NAVY
                  ).grid(row=0, column=1, sticky="w")
        self.orig_label = ttk.Label(edit, textvariable=self.sel_orig_var,
                                    foreground="#666666", width=30)
        self.orig_label.grid(row=0, column=2, columnspan=5, sticky="ew", padx=(4, 0))

        ttk.Label(edit, text="Neu:").grid(row=1, column=0, sticky="nw", pady=(4, 0))
        self.new_entry = VarText(edit, self.sel_new_var, height=3)
        self.new_entry.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(4, 6), pady=(4, 0))
        self.sel_new_var.trace_add("write", self._on_new_value_typed)

        self.clear_cb = ttk.Checkbutton(edit, text="Feld leeren",
                                        variable=self.sel_clear_var,
                                        command=self._on_clear_toggled)
        self.clear_cb.grid(row=1, column=4, sticky="n", pady=(4, 0))
        ToolTip(self.clear_cb, "Den Inhalt des gewählten Feldes in allen "
                               "abgelegten Spezifikationen entfernen")
        ttk.Button(edit, text="Zurücknehmen", width=13,
                   command=self._revert_current).grid(row=1, column=5, sticky="n", padx=(6, 0), pady=(4, 0))

        self.adopt_cb = ttk.Checkbutton(
            edit, text="Vorlagenwert in alle Spezifikationen übernehmen",
            variable=self.adopt_var, command=self._on_adopt_toggled)
        self.adopt_cb.grid(row=2, column=1, columnspan=5, sticky="w",
                           padx=(4, 0), pady=(4, 0))
        ToolTip(self.adopt_cb,
                "Der aktuelle Feldwert der Vorlage (z. B. Revision 003) wird in "
                "jeder hereingezogenen Spezifikation gesetzt. Fehlt das Feld "
                "dort noch, wird es an der passenden Stelle eingefügt.")

        self._edit_widgets = (self.new_entry, self.clear_cb, self.adopt_cb)
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
        self.sel_new_var.set(orig if cell.new_text is None else cell.new_text)
        self.sel_clear_var.set(cell.new_text == "")
        self.adopt_var.set(cell.new_text is not None and cell.new_text == orig)
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
        same = text == self.sheet.orig_value(coord)
        keep = same and self.adopt_var.get()
        unchanged = not text or (same and not keep)
        self.sheet.set_new_value(coord, None if unchanged else text)
        self._refresh_counter()

    def _on_adopt_toggled(self):
        coord = self._current_coord()
        if not coord:
            return
        orig = self.sheet.orig_value(coord)
        if self.adopt_var.get():
            self._edit_guard = True
            self.sel_clear_var.set(False)
            self.sel_new_var.set(orig)
            self._edit_guard = False
            self.new_entry.configure(state="normal")
            self.sheet.set_new_value(coord, orig)
        else:
            self.sheet.set_new_value(coord, None)
        self._refresh_counter()

    def _on_clear_toggled(self):
        coord = self._current_coord()
        if not coord:
            return
        if self.sel_clear_var.get():
            self._edit_guard = True
            self.adopt_var.set(False)
            self.sel_new_var.set("")
            self._edit_guard = False
            self.new_entry.configure(state="disabled")
            self.sheet.set_new_value(coord, "")
        else:
            self.new_entry.configure(state="normal")
            self._edit_guard = True
            self.sel_new_var.set(self.sheet.orig_value(coord))
            self._edit_guard = False
            self.sheet.set_new_value(coord, None)
        self._refresh_counter()

    def _revert_current(self):
        coord = self._current_coord()
        if not coord:
            return
        self._edit_guard = True
        self.sel_new_var.set(self.sheet.orig_value(coord))
        self.sel_clear_var.set(False)
        self.adopt_var.set(False)
        self._edit_guard = False
        self.new_entry.configure(state="normal")
        self.sheet.set_new_value(coord, None)
        self._refresh_counter()

    def _clear_all_changes(self):
        self.sheet.clear_all()
        self._edit_guard = True
        self.sel_new_var.set(self.sheet.orig_value(self._selected_coord or ""))
        self.sel_clear_var.set(False)
        self.adopt_var.set(False)
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
        frame.grid(row=0, column=0, sticky="ew", pady=(2, 6))
        frame.columnconfigure(0, weight=1)

        ttk.Radiobutton(frame, text="Zielordner:", variable=self.out_mode,
                        value="folder").grid(row=0, column=0, columnspan=2, sticky="w")
        dir_entry = ttk.Entry(frame, textvariable=self.out_dir, width=30)
        dir_entry.grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(frame, text="Durchsuchen…",
                   command=self._browse_dir).grid(row=1, column=1)

        ttk.Radiobutton(frame, text="Nach jeder Datei fragen („Speichern unter“)",
                        variable=self.out_mode,
                        value="ask").grid(row=2, column=0, columnspan=2,
                                          sticky="w", pady=(6, 0))

        suffix_row = ttk.Frame(frame)
        suffix_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(suffix_row, text="Dateinamen-Zusatz:").pack(side="left")
        suffix_entry = ttk.Entry(suffix_row, textvariable=self.suffix, width=14)
        suffix_entry.pack(side="left", padx=6)
        ToolTip(suffix_entry, "Wird an den Dateinamen angehängt, z. B. "
                              "„Spezi_angepasst.xlsx“.\nLeer lassen = Originalname.")

        opt = ttk.Frame(frame)
        opt.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Label(opt, text="Ersetzungsmodus:").pack(anchor="w")
        rb1 = ttk.Radiobutton(opt, text="Ganze Zelle ersetzen",
                              variable=self.replace_mode, value="set")
        rb1.pack(anchor="w", padx=(8, 0))
        rb2 = ttk.Radiobutton(opt, text="Nur Teiltext ersetzen",
                              variable=self.replace_mode, value="replace")
        rb2.pack(anchor="w", padx=(8, 0))
        ToolTip(rb2, "Ersetzt den aktuellen Feldtext innerhalb längerer "
                     "Zelltexte (z. B. Datum im Kopftext).")
        verify_cb = ttk.Checkbutton(
            opt, text="Nur bei gleichem Vorlagenwert ersetzen",
            variable=self.verify)
        verify_cb.pack(anchor="w", pady=(6, 0))
        ToolTip(verify_cb, "Sicherheitsoption: Zellen, deren Inhalt in der "
                           "abgelegten Spezifikation vom Vorlagenwert abweicht, "
                           "werden übersprungen und im Log gemeldet. Leere "
                           "Eingabefelder werden nur befüllt, wenn sie in der "
                           "Zieldatei leer sind.")
        layout_cb = ttk.Checkbutton(
            opt, text="Layout-Prüfung (empfohlen)", variable=self.layout_check)
        layout_cb.pack(anchor="w", pady=(2, 0))
        ToolTip(layout_cb, "Prüft jede abgelegte Datei vor dem Ersetzen: Weicht "
                           "ihr Layout deutlich von der Vorlage ab (z. B. andere "
                           "Zeilenanzahl), wird sie übersprungen statt falsch "
                           "befüllt - mit Warnung im Log. Verschobene Zeilen "
                           "und geänderte Werte werden toleriert.")

    # ------------------------------------------------------------------
    # Schritt 3: Verarbeiten
    # ------------------------------------------------------------------

    def _build_step3(self, main):
        frame = ttk.LabelFrame(
            main, text=" 3 · Spezifikationen verarbeiten ", padding=8)
        frame.grid(row=1, column=0, sticky="nsew", pady=(0, 2))
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
        self.log = ttk.Treeview(log_frame, columns=cols, show="headings", height=4)
        self.log.heading("file", text="Datei")
        self.log.heading("result", text="Ergebnis")
        self.log.heading("target", text="Ziel")
        self.log.column("file", width=120, anchor="w")
        self.log.column("result", width=170, anchor="w")
        self.log.column("target", width=110, anchor="w")
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
        self.progress.grid_remove()
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
        line1 = "⬇  Dateien hierher ziehen (.xlsx)"
        if not self.dnd_ok:
            line1 = "Drag & Drop nicht verfügbar – bitte die Schaltfläche unten nutzen"
        c.create_text(w / 2, h / 2 - 10, text=line1,
                      font=("Segoe UI", 11, "bold"), fill=self._canvas_fg)
        c.create_text(w / 2, h / 2 + 14,
                      text="Mehrfachauswahl möglich",
                      font=("Segoe UI", 9), fill=self._canvas_muted)

    # ------------------------------------------------------------------
    # Vorlage / Sprache
    # ------------------------------------------------------------------

    def _show_placeholder(self, show: bool):
        if show:
            self.placeholder.grid(row=1, column=0, sticky="nsew")
            self.placeholder.lift()
        else:
            self.placeholder.grid_remove()
        self._set_edit_enabled(False)

    def _pick_template(self):
        path = filedialog.askopenfilename(
            parent=self.root, title="Vorlage auswählen (korrekte Spezifikation)",
            filetypes=FILE_TYPES)
        if path:
            self.set_template(path)

    def _on_template_drop(self, event):
        paths = [p for p in self.root.tk.splitlist(event.data)
                 if p.lower().endswith((".xlsx", ".xlsm"))]
        if paths:
            self.set_template(paths[0])

    def set_template(self, path: str) -> bool:
        """Lädt eine neue Vorlage; bisherige Markierungen werden verworfen."""
        previous = self.template_path
        self.template_path = path
        if not self._load_template(preserve={}):
            self.template_path = previous
            return False
        self._save_cfg()
        return True

    def _zoom_fit(self) -> bool:
        return not self.zoom_var.get().strip().endswith("%")

    def _zoom_scale(self) -> float:
        try:
            return max(40, min(200, int(self.zoom_var.get().rstrip(" %")))) / 100
        except ValueError:
            return 1.0

    def _on_zoom_change(self, _event=None):
        self._load_template(preserve=self.sheet.changed())
        self._save_cfg()

    def _load_template(self, preserve: dict) -> bool:
        if not self.template_path:
            return False
        self.status_var.set("Analysiere Vorlage…")
        self.root.update_idletasks()
        try:
            count = self.sheet.load_template(self.template_path, keep=preserve,
                                             scale=self._zoom_scale(),
                                             fit=self._zoom_fit(),
                                             page=self.sheet.page)
            self._layout_ref = layout_fingerprints(self.template_path)
            self._row_texts = template_row_texts(self.template_path)
        except Exception as exc:
            messagebox.showerror(APP_NAME,
                                 f"Vorlage konnte nicht gelesen werden:\n{exc}")
            self.status_var.set("Vorlage konnte nicht gelesen werden.")
            return False
        self._show_placeholder(False)
        self.template_var.set(os.path.basename(self.template_path))
        self._refresh_counter()
        self.search_info_var.set("")
        self.search_var.set("")
        self.status_var.set(
            f"Vorlage: {count} Felder erkannt – bitte Felder anklicken, "
            f"die in alle Spezifikationen übernommen werden sollen.")
        return True

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
            find, repl, rmode = orig, new, mode
            line_start, line_count = -1, 0
            if not coord.startswith(HF_PREFIX) and new and "\n" in orig:
                part = _changed_lines(orig, new)
                if part:
                    find, repl, rmode = part[0], part[1], "replace"
                    line_start, line_count = part[2], part[3]
            is_hf = coord.startswith(HF_PREFIX)
            rules.append(Rule(
                coord=coord, new=repl, mode=rmode,
                find=find, expected=orig,
                line_start=line_start, line_count=line_count,
                label=label,
                anchor_text=anchor_text, anchor_coord=anchor_coord,
                context_above=([] if is_hf else context_above(
                    self._row_texts, split_coord(coord)[1])),
                allow_insert=not is_hf,
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

        if not self.template_path:
            messagebox.showwarning(
                APP_NAME, "Bitte zuerst oben eine Vorlage hochladen.")
            self.status_var.set("Keine Vorlage geladen – nichts zu tun.")
            return

        rules = self._collect_rules()
        if not rules:
            messagebox.showwarning(
                APP_NAME,
                "Bitte zuerst in Schritt 1 mindestens ein Feld anklicken\n"
                "und „in alle Spezifikationen übernehmen“ wählen\n"
                "(oder einen neuen Wert eintragen).")
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

        self._start_progress(len(files))
        self.open_btn.configure(state="disabled")
        self.status_var.set(f"Verarbeite {len(files)} Datei(en)…")

        self._worker = threading.Thread(
            target=self._batch_worker,
            args=(files, rules, out_dir, self.suffix.get(), self.verify.get(),
                  self.layout_check.get(), self._layout_ref_for(rules)),
            daemon=True)
        self._worker.start()

    def _layout_ref_for(self, rules: list[Rule]) -> dict:
        """Vorlagen-Fingerabdruck ohne die Zellen, die ohnehin geändert werden."""
        changed = {r.coord for r in rules}
        return {k: v for k, v in self._layout_ref.items()
                if k.split("!", 1)[-1] not in changed or not k.startswith("0!")}

    def _batch_worker(self, files, rules, out_dir, suffix, verify,
                      layout_check, layout_ref):
        errors = []
        for i, src in enumerate(files, start=1):
            name = os.path.basename(src)
            self._queue.put(("file_start", i, name))
            try:
                if layout_check and layout_ref:
                    sim = layout_similarity(layout_ref, layout_fingerprints(src))
                    if sim < LAYOUT_MIN_SIMILARITY:
                        self._queue.put(("log", name, "warn",
                                         f"Layout weicht ab (nur {sim:.0%} Übereinstimmung"
                                         f" mit der Vorlage) – übersprungen. "
                                         f"Tipp: Layout-Prüfung in Schritt 2 "
                                         f"abschalten, falls die Datei passt.", ""))
                    else:
                        self._apply_one(src, out_dir, suffix, rules, verify, errors)
                else:
                    self._apply_one(src, out_dir, suffix, rules, verify, errors)
            except Exception as exc:
                errors.append(f"{name}: {exc}")
                self._queue.put(("log", name, "err", f"Fehler: {exc}", ""))
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
        if self._worker and self._worker.is_alive():
            messagebox.showinfo(APP_NAME, "Es läuft bereits eine Verarbeitung.")
            return

        errors = []
        layout_check = self.layout_check.get()
        layout_ref = self._layout_ref_for(rules)
        verify = self.verify.get()
        self._start_progress(len(files))
        self.open_btn.configure(state="disabled")
        self.status_var.set(f"Verarbeite {len(files)} Datei(en)…")
        for i, src in enumerate(files, start=1):
            name = os.path.basename(src)
            self._progress_current = name
            if layout_check and self._layout_ref:
                try:
                    sim = self._run_dialog_step(
                        lambda: layout_similarity(
                            layout_ref, layout_fingerprints(src)))
                except Exception as exc:
                    errors.append(f"{name}: {exc}")
                    self._log(name, "err", f"Fehler: {exc}", "")
                    self._mark_progress_done(i)
                    continue
                if sim < LAYOUT_MIN_SIMILARITY and not messagebox.askyesno(
                        APP_NAME,
                        f"{name}\n\nLayout weicht ab (nur {sim:.0%} Übereinstimmung "
                        f"mit der Vorlage).\n\nTrotzdem verarbeiten?",
                        parent=self.root):
                    self._log(name, "warn",
                              f"Layout weicht ab ({sim:.0%}) – übersprungen", "")
                    self._mark_progress_done(i)
                    continue
            stem, ext = os.path.splitext(name)
            target = filedialog.asksaveasfilename(
                parent=self.root, title=f"Geänderte Spezifikation speichern: {name}",
                initialfile=f"{stem}{self.suffix.get()}{ext}",
                defaultextension=ext, filetypes=FILE_TYPES)
            if not target:
                self._log(name, "warn", "Speichern abgebrochen", "")
                self._mark_progress_done(i)
                continue
            try:
                res = self._run_dialog_step(
                    lambda: apply_to_file(
                        src, target, rules, verify=verify))
            except Exception as exc:
                errors.append(f"{name}: {exc}")
                self._log(name, "err", f"Fehler: {exc}", "")
                self._mark_progress_done(i)
                continue
            if res.ok:
                self._log(name, "ok", f"{res.changed_count} Feld(er) geändert", target)
                self._last_out_dir = os.path.dirname(target)
            else:
                errors.append(f"{name}: {res.error}")
                self._log(name, "err", f"Fehler: {res.error}", "")
            self._mark_progress_done(i)
        self.status_var.set(f"Fertig – {len(files)} Datei(en) verarbeitet.")
        self._finish_progress(len(files))
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
                    self._progress_done = item[1]
                elif kind == "file_start":
                    self._progress_current = item[2]
                elif kind == "done":
                    _, total, errors = item
                    self._finish_progress(total)
                    if errors:
                        messagebox.showerror(
                            APP_NAME, "Fehler bei:\n\n" + "\n".join(errors))
        except queue.Empty:
            pass
        if (self._progress_visible and self._progress_started_at is not None
                and self._worker and self._worker.is_alive()):
            self._update_progress_status()
        self.root.after(60, self._poll_queue)

    def _start_progress(self, total: int):
        if self._progress_reveal_job is not None:
            self.root.after_cancel(self._progress_reveal_job)
            self._progress_reveal_job = None
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=total)
        self.progress.grid_remove()
        self.progress_var.set(0)
        self._progress_started_at = time.monotonic()
        self._progress_total = total
        self._progress_done = 0
        self._progress_current = ""
        self._progress_visible = False
        self._progress_reveal_job = self.root.after(
            700, self._reveal_progress)

    def _finish_progress(self, total: int):
        if self._progress_reveal_job is not None:
            self.root.after_cancel(self._progress_reveal_job)
            self._progress_reveal_job = None
        self._progress_started_at = None
        self.progress.stop()
        self.progress.configure(mode="determinate", maximum=total)
        self.progress_var.set(total)
        self.status_var.set(f"Fertig – {total} Datei(en) verarbeitet.")
        self.open_btn.configure(state="normal")

    def _mark_progress_done(self, count: int):
        self._progress_done = count
        self.progress_var.set(count)

    def _run_dialog_step(self, callback):
        result = []
        error = []
        finished = threading.Event()

        def run():
            try:
                result.append(callback())
            except Exception as exc:
                error.append(exc)
            finally:
                finished.set()

        self._worker = threading.Thread(target=run, daemon=True)
        self._worker.start()
        wait_var = tk.BooleanVar(value=False)

        def check_finished():
            if finished.is_set():
                wait_var.set(True)
            else:
                self.root.after(30, check_finished)

        self.root.after(30, check_finished)
        self.root.wait_variable(wait_var)
        self._worker.join()
        self._worker = None
        if error:
            raise error[0]
        return result[0]

    def _reveal_progress(self):
        self._progress_reveal_job = None
        if self._progress_started_at is None:
            return
        self._progress_visible = True
        self.progress.grid()
        if self._progress_total == 1:
            self.progress.configure(mode="indeterminate")
            self.progress.start(12)
        self._update_progress_status()

    def _update_progress_status(self):
        if self._progress_started_at is None:
            return
        elapsed = time.monotonic() - self._progress_started_at
        elapsed_text = self._format_duration(elapsed)
        current = self._progress_current or "Datei wird vorbereitet"
        status = (f"{current} · {self._progress_done}/{self._progress_total} "
                  f"fertig · läuft seit {elapsed_text}")
        if self._progress_done:
            remaining = elapsed / self._progress_done * (
                self._progress_total - self._progress_done)
            status += f" · noch ca. {self._format_duration(remaining)}"
        self.status_var.set(status)

    @staticmethod
    def _format_duration(seconds: float) -> str:
        total_seconds = max(0, int(seconds))
        minutes, seconds = divmod(total_seconds, 60)
        if minutes:
            return f"{minutes}:{seconds:02d} min"
        return f"{seconds} s"

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
