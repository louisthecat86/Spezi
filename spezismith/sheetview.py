"""
SpeziSmith - SheetView

Maßstabsgetreue, klickbare Nachbildung der Spezifikations-Vorlage als
**ganze Seite**: Seitenkopf und -fuß (Kopf-/Fußzeile) werden wie im
Ausdruck oben bzw. unten dargestellt und sind ebenfalls wählbar.
Spaltenbreiten, Zeilenhöhen, verbundene Zellen, Schrift- und Füllfarben
kommen aus der Excel-Datei; Seitenumbrüche werden als gestrichelte Linien
eingeblendet. Zoombar. Ein Klick auf ein Feld wählt es aus, eingetragene
Änderungen werden direkt im Layout sichtbar (WYSIWYG).
"""

from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
from typing import Callable, Optional

import openpyxl

from .engine import HF_PREFIX, hf_display_text, idx_to_col

ORANGE = "#F08C21"
NAVY = "#1F2942"
CHANGED_BG = "#FFE3C2"
MATCH_BG = "#FFF59D"
HEADER_BG = "#E9EDF3"
GRID_COLOR = "#B9BFC9"
HF_BG = "#F7F9FC"

ROW_HEADER_W = 44
COL_HEADER_H = 22
DEFAULT_COL_WIDTH = 8.43   # Excel-Standardbreite
DEFAULT_ROW_HEIGHT = 16.5  # Punkte

HF_PART_NAMES = {
    "oddHeader": "Kopfzeile", "oddFooter": "Fußzeile",
    "evenHeader": "Kopfzeile (gerade Seiten)",
    "evenFooter": "Fußzeile (gerade Seiten)",
    "firstHeader": "Kopfzeile (1. Seite)",
    "firstFooter": "Fußzeile (1. Seite)",
}
HF_SECTION_NAMES = {"L": "links", "C": "mittig", "R": "rechts"}


def hf_friendly_name(coord: str) -> str:
    """'HF:oddFooter.R' -> 'Fußzeile · rechts'"""
    try:
        part, section = coord[len(HF_PREFIX):].rsplit(".", 1)
        return f"{HF_PART_NAMES.get(part, part)} · {HF_SECTION_NAMES.get(section, section)}"
    except ValueError:
        return coord


def _px_width(excel_width: float, scale: float) -> int:
    return max(8, int((excel_width * 7 + 6) * scale))


def _px_height(points: float, scale: float) -> int:
    return max(3, int(points * 96 / 72 * scale))


def _excel_color(color_obj, fallback: str) -> str:
    try:
        if color_obj is not None and color_obj.type == "rgb":
            rgb = color_obj.rgb
            if isinstance(rgb, str) and len(rgb) >= 6 and rgb != "00000000":
                return "#" + rgb[-6:]
    except Exception:
        pass
    return fallback


class CellWidget:
    """Ein klickbares Feld im Seitenraster (Zelle oder Kopf-/Fußzeilen-Teil)."""

    def __init__(self, sheet: "SheetView", coord: str, row: int, col: int,
                 grid_row: int, text: str, w_px: int, h_px: int,
                 rowspan: int, colspan: int,
                 font: tkfont.Font, bg: str, fg: str, anchor: str,
                 justify: str, wrap: bool):
        self.sheet = sheet
        self.coord = coord
        self.grid_row = grid_row
        self.row = row
        self.col = col
        self.orig_text = text
        self.new_text: Optional[str] = None   # None = unverändert
        self.base_bg = bg

        self.frame = tk.Frame(
            sheet.grid_frame, width=w_px, height=h_px,
            background=GRID_COLOR, borderwidth=0, highlightthickness=0,
            highlightbackground=ORANGE, highlightcolor=ORANGE)
        self.frame.grid_propagate(False)
        self.frame.pack_propagate(False)

        self.label = tk.Label(
            self.frame, text=text, font=font, fg=fg, bg=bg,
            anchor=anchor, justify=justify, padx=3, pady=0, borderwidth=0,
            highlightthickness=1, highlightbackground=GRID_COLOR,
            wraplength=max(10, w_px - 8) if wrap else 0)
        self.label.pack(fill="both", expand=True)

        self.frame.grid(row=grid_row, column=col, rowspan=rowspan,
                        columnspan=colspan, sticky="nsew")
        for w in (self.frame, self.label):
            w.bind("<Button-1>", lambda _e, c=coord: sheet.select(c))
        tip = f"{coord}: {text or '(leer)'}"
        if coord.startswith(HF_PREFIX):
            tip = f"{hf_friendly_name(coord)}: {text}\n" \
                  "(&P = Seitenzahl, &N = Seiten gesamt, &G = Grafik)"
        sheet.register_tooltip(self.label, tip)

    # -- Zustand ------------------------------------------------------

    @property
    def changed(self) -> bool:
        return self.new_text is not None

    def set_new_text(self, text: Optional[str]):
        self.new_text = text
        shown = self.orig_text if text is None else text
        self.label.configure(text=shown)
        self._refresh()

    def _refresh(self):
        selected = self.sheet.selected == self.coord
        bg = CHANGED_BG if self.changed else self.base_bg
        self.label.configure(background=bg)
        self.frame.configure(highlightthickness=2 if selected else 0)

    def flash_match(self, on: bool):
        self.label.configure(
            background=MATCH_BG if on else (CHANGED_BG if self.changed else self.base_bg))


class SheetView(ttk.Frame):
    """Scrollbare, klickbare Seitenansicht mit fixierten Kopfleisten."""

    def __init__(self, master, on_select: Callable[[Optional[str]], None]):
        super().__init__(master)
        self.on_select = on_select
        self.cells: dict[str, CellWidget] = {}
        self.anchors: dict[str, tuple[str, str]] = {}
        self.selected: Optional[str] = None
        self.scale = 1.0
        self._fonts: dict[tuple, tkfont.Font] = {}
        self._matches: list[str] = []
        self._match_idx = -1
        self.fit = False
        self.pages: list[list[int]] = [[]]
        self.page = 0
        self._row_items: dict[int, list] = {}
        self._row_min: dict[int, tuple[int, int]] = {}
        self._path: Optional[str] = None
        self._fit_job = None

        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight=1)

        # Eckbox + Kopfleisten (fixiert, werden synchron gescrollt)
        self.corner = tk.Label(self, text="", background=HEADER_BG,
                               relief="solid", borderwidth=1)
        self.corner.grid(row=0, column=0, sticky="nsew")
        self.top_canvas = tk.Canvas(self, height=COL_HEADER_H,
                                    highlightthickness=0, background=HEADER_BG)
        self.top_canvas.grid(row=0, column=1, sticky="ew")
        self.left_canvas = tk.Canvas(self, width=ROW_HEADER_W,
                                     highlightthickness=0, background=HEADER_BG)
        self.left_canvas.grid(row=1, column=0, sticky="ns")

        # Hauptfläche
        self.canvas = tk.Canvas(self, highlightthickness=0, background="white")
        self.canvas.grid(row=1, column=1, sticky="nsew")
        vscroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        vscroll.grid(row=1, column=2, sticky="ns")
        hscroll = ttk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        hscroll.grid(row=2, column=1, sticky="ew")
        self.canvas.configure(
            xscrollcommand=lambda *a: (hscroll.set(*a),
                                       self.top_canvas.xview_moveto(a[0])),
            yscrollcommand=lambda *a: (vscroll.set(*a),
                                       self.left_canvas.yview_moveto(a[0])),
        )

        # Seitennavigation (unter der Seitenansicht)
        self.nav = ttk.Frame(self)
        self.nav.grid(row=3, column=0, columnspan=3, pady=(4, 0))
        self.prev_btn = ttk.Button(self.nav, text="◀", width=3,
                                   command=lambda: self.show_page(self.page - 1))
        self.prev_btn.pack(side="left")
        self.page_label = ttk.Label(self.nav, text="Seite 1 von 1",
                                    font=("Segoe UI", 10, "bold"))
        self.page_label.pack(side="left", padx=10)
        self.next_btn = ttk.Button(self.nav, text="▶", width=3,
                                   command=lambda: self.show_page(self.page + 1))
        self.next_btn.pack(side="left")
        self.jump_frame = ttk.Frame(self.nav)
        self.jump_frame.pack(side="left", padx=(16, 0))
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        self.grid_frame = tk.Frame(self.canvas, background="white")
        self.top_frame = tk.Frame(self.top_canvas, background=HEADER_BG)
        self.left_frame = tk.Frame(self.left_canvas, background=HEADER_BG)
        self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")
        self.top_canvas.create_window((0, 0), window=self.top_frame, anchor="nw")
        self.left_canvas.create_window((0, 0), window=self.left_frame, anchor="nw")
        self.grid_frame.bind("<Configure>", lambda _e: self._update_scrollregions())
        self.top_frame.bind("<Configure>", lambda _e: self._update_scrollregions())
        self.left_frame.bind("<Configure>", lambda _e: self._update_scrollregions())

        for w in (self.canvas, self.grid_frame):
            w.bind("<MouseWheel>", self._on_wheel)
            w.bind("<Button-4>", self._on_wheel)
            w.bind("<Button-5>", self._on_wheel)

        self._tooltips: list = []

    # -- Aufbau ---------------------------------------------------------

    def register_tooltip(self, widget, text):
        from .gui import ToolTip  # zirkulären Import vermeiden
        self._tooltips.append(ToolTip(widget, text))

    def _font(self, bold: bool, size: float, italic: bool = False) -> tkfont.Font:
        key = (bold, round(size * self.scale, 1), italic)
        if key not in self._fonts:
            self._fonts[key] = tkfont.Font(
                family="Segoe UI", size=max(5, int(round(size * self.scale))),
                weight="bold" if bold else "normal",
                slant="italic" if italic else "roman")
        return self._fonts[key]

    def load_template(self, path: str, keep: Optional[dict] = None,
                      scale: Optional[float] = None, fit: bool = False,
                      page: Optional[int] = None) -> int:
        """Lädt eine Vorlage. fit=True: Maßstab so wählen, dass eine ganze
        Seite ohne Scrollen in die Ansicht passt."""
        self._path = path
        self.fit = fit
        if scale is not None and not fit:
            self.scale = scale
        if page is not None:
            self.page = page
        count = self._build(path, keep)
        if fit:
            self._autofit(keep)
        return count

    def _fit_factor(self) -> Optional[float]:
        """Faktor, um den der Maßstab geändert werden müsste (1.0 = passt exakt)."""
        self.update_idletasks()
        avail_w, avail_h = self.canvas.winfo_width(), self.canvas.winfo_height()
        if avail_w < 100 or avail_h < 100:
            return None
        need_w = need_h = 1
        current = self.page
        for n in range(len(self.pages)):
            self._apply_page(n)
            self.grid_frame.update_idletasks()
            need_w = max(need_w, self.grid_frame.winfo_reqwidth())
            need_h = max(need_h, self.grid_frame.winfo_reqheight())
        self._apply_page(current)
        return min((avail_w - 4) / need_w, (avail_h - 4) / need_h)

    def _autofit(self, keep: Optional[dict]):
        for _ in range(10):
            factor = self._fit_factor()
            if factor is None or 1.0 <= factor <= 1.05:
                break
            # Feste Mindestgrößen machen die Höhe nicht linear: beim Verkleinern
            # stärker nachregeln, beim Vergrößern vorsichtig herantasten.
            step = factor ** 3 * 0.99 if factor < 1.0 else 1.0 + (factor - 1.0) * 0.5
            new = max(0.3, min(1.5, self.scale * step))
            if abs(new - self.scale) < 0.005:
                break
            self.scale = new
            self._build(self._path, keep if keep is not None else self.changed())
        self._apply_page(self.page)

    def _on_canvas_resize(self, _event=None):
        if not self.fit or not self._path:
            return
        if self._fit_job:
            self.after_cancel(self._fit_job)
        self._fit_job = self.after(250, self._refit)

    def _refit(self):
        self._fit_job = None
        factor = self._fit_factor()
        if factor is not None and not 1.0 <= factor <= 1.05:
            keep = self.changed()
            selected = self.selected
            self._autofit(keep)
            if selected in self.cells:
                self.selected = selected
                self.cells[selected]._refresh()

    def _build(self, path: str, keep: Optional[dict] = None) -> int:
        """Baut die komplette Seite neu auf."""
        for cell in self.cells.values():
            cell.frame.destroy()
        for w in self.top_frame.winfo_children() + self.left_frame.winfo_children():
            w.destroy()
        for w in self.grid_frame.winfo_children():
            w.destroy()
        self.cells.clear()
        self._row_items.clear()
        self._row_min.clear()
        self.anchors.clear()
        self._tooltips.clear()
        self.selected = None
        self._matches = []
        keep = keep or {}
        sc = self.scale

        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            ws = wb.worksheets[0]
            max_row, max_col = ws.max_row, ws.max_column

            col_px = [ROW_HEADER_W]
            for c in range(1, max_col + 1):
                dim = ws.column_dimensions.get(idx_to_col(c))
                width = dim.width if (dim and dim.width) else DEFAULT_COL_WIDTH
                col_px.append(_px_width(width, sc))
            row_px = [COL_HEADER_H]
            for r in range(1, max_row + 1):
                dim = ws.row_dimensions.get(r)
                height = dim.height if (dim and dim.height) else DEFAULT_ROW_HEIGHT
                row_px.append(_px_height(height, sc))

            # Kopf-/Fußzeile (wie im Ausdruck)
            hf_defs = (
                ("top", "oddHeader", ws.oddHeader),
                ("top", "firstHeader", ws.firstHeader),
                ("top", "evenHeader", ws.evenHeader),
                ("bottom", "oddFooter", ws.oddFooter),
                ("bottom", "firstFooter", ws.firstFooter),
                ("bottom", "evenFooter", ws.evenFooter),
            )
            top_parts: list[tuple[str, str]] = []
            bottom_parts: list[tuple[str, str]] = []
            for where, part, obj in hf_defs:
                for section, item in (("L", obj.left), ("C", obj.center),
                                      ("R", obj.right)):
                    if item.text:
                        entry = (f"{HF_PREFIX}{part}.{section}",
                                 hf_display_text(item.text).strip("\n"))
                        (top_parts if where == "top" else bottom_parts).append(entry)

            # Seitenumbrüche (gestrichelte Linien) - nicht mitten durch Verbund
            breaks = set()
            if ws.row_breaks and ws.row_breaks.brk:
                breaks = {int(b.id) for b in ws.row_breaks.brk if b.id}
            for rng in ws.merged_cells.ranges:
                for br in list(breaks):
                    if rng.min_row <= br < rng.max_row:
                        breaks.discard(br)

            # Seiten-Plan: Kopfzeile(n), Datenzeilen + Umbrüche, Fußzeile(n)
            plan: list[tuple[str, object]] = []
            if top_parts:
                plan.append(("hf_top", top_parts))
            for r in range(1, max_row + 1):
                plan.append(("row", r))
                if r in breaks:
                    plan.append(("sep", r))
            if bottom_parts:
                plan.append(("hf_bottom", bottom_parts))

            self.pages = [[]]
            for idx, (kind, _payload) in enumerate(plan):
                if kind == "sep":
                    self.pages.append([])
                else:
                    self.pages[-1].append(idx)
            if len(self.pages) > 1 and not self.pages[-1]:
                self.pages.pop()
            self.page = max(0, min(self.page, len(self.pages) - 1))

            grid_row_of: dict[int, int] = {}
            for idx, (kind, payload) in enumerate(plan):
                if kind == "row":
                    grid_row_of[payload] = idx

            # Kopfleiste (Spaltenbuchstaben)
            for c in range(1, max_col + 1):
                tk.Label(self.top_frame, text=idx_to_col(c), background=HEADER_BG,
                         foreground="#4A5568", font=self._font(True, 9),
                         relief="solid", borderwidth=1,
                         ).grid(row=0, column=c, sticky="nsew", ipady=2)
                self.top_frame.grid_columnconfigure(c, minsize=col_px[c])
            for c in range(1, max_col + 1):
                self.grid_frame.grid_columnconfigure(c, minsize=col_px[c])

            hf_font = self._font(False, 8.5, italic=True)

            def hf_row_height(parts):
                lines = max((txt.count("\n") + 1 for _c, txt in parts), default=1)
                return int((lines * 13 + 10) * sc) + 6

            # Zeilen nach Plan aufbauen
            for idx, (kind, payload) in enumerate(plan):
                if kind == "row":
                    r = payload
                    self.grid_frame.grid_rowconfigure(idx, minsize=row_px[r])
                    lbl = self._row_label(str(r), row_px[r], "#4A5568",
                                          self._font(True, 9), idx)
                    self.left_frame.grid_rowconfigure(idx, minsize=row_px[r])
                    self._row_items.setdefault(idx, []).append(lbl)
                    self._row_min[idx] = (row_px[r], row_px[r])

                elif kind == "sep":
                    self.grid_frame.grid_rowconfigure(idx, minsize=8)
                    sep = tk.Canvas(self.grid_frame, height=8, highlightthickness=0,
                                    background="white")
                    sep.grid(row=idx, column=1, columnspan=max_col, sticky="ew")
                    sep.bind("<Configure>", lambda e, cv=sep: self._draw_sep(cv))
                    lbl = self._row_label("", 8, "#4A5568", self._font(True, 9), idx)
                    self.left_frame.grid_rowconfigure(idx, minsize=8)
                    self._row_items.setdefault(idx, []).extend([lbl, sep])
                    self._row_min[idx] = (8, 8)

                elif kind in ("hf_top", "hf_bottom"):
                    parts = payload
                    h_px = hf_row_height(parts)
                    self.grid_frame.grid_rowconfigure(idx, minsize=h_px)
                    third = max(1, max_col // 3)
                    spans = ((1, third), (third + 1, 2 * third),
                             (2 * third + 1, max_col))
                    for (coord, text), (c1, c2) in zip(parts, spans):
                        section = coord.rsplit(".", 1)[1]
                        anchor = {"L": "w", "C": "center", "R": "e"}[section]
                        justify = {"L": "left", "C": "center", "R": "right"}[section]
                        w_px = sum(col_px[c1:c2 + 1]) + (c2 - c1)
                        widget = CellWidget(
                            self, coord, -1, c1, idx, text, w_px, h_px,
                            1, c2 - c1 + 1, hf_font, HF_BG, "#333333",
                            anchor, justify, True)
                        self.cells[coord] = widget
                        self._row_items.setdefault(idx, []).append(widget.frame)
                        if coord in keep:
                            widget.set_new_text(keep[coord])
                    label = "Kopf" if kind == "hf_top" else "Fuß"
                    lbl = self._row_label(label, h_px, "#98A2B3",
                                          self._font(False, 7), idx)
                    self.left_frame.grid_rowconfigure(idx, minsize=h_px)
                    self._row_items.setdefault(idx, []).append(lbl)
                    self._row_min[idx] = (h_px, h_px)

            # Verbundene Zellen
            span: dict[str, tuple[int, int]] = {}
            covered: set[str] = set()
            for rng in ws.merged_cells.ranges:
                anchor = f"{idx_to_col(rng.min_col)}{rng.min_row}"
                span[anchor] = (rng.min_row, rng.max_row, rng.min_col, rng.max_col)
                for r in range(rng.min_row, rng.max_row + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        if (r, c) != (rng.min_row, rng.min_col):
                            covered.add(f"{idx_to_col(c)}{r}")

            # Daten-Zellen
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    coord = f"{idx_to_col(c)}{r}"
                    if coord in covered:
                        continue
                    cell = ws.cell(row=r, column=c)
                    r1, r2, c1, c2 = span.get(coord, (r, r, c, c))
                    rs = grid_row_of[r2] - grid_row_of[r1] + 1
                    cs = c2 - c1 + 1
                    w_px = sum(col_px[c1:c2 + 1]) + (cs - 1)
                    h_px = sum(row_px[r1:r2 + 1]) + (rs - 1)

                    text = self._format_value(cell.value)
                    bold = bool(cell.font and cell.font.b)
                    size = float(cell.font.sz) if (cell.font and cell.font.sz) else 11.0
                    fg = _excel_color(cell.font.color if cell.font else None, "#1a1a1a")
                    fill = cell.fill
                    bg = _excel_color(fill.fgColor, "#FFFFFF") \
                        if (fill and fill.patternType) else "#FFFFFF"
                    align = cell.alignment
                    horizontal = align.horizontal if align else None
                    anchor = {"left": "w", "center": "center",
                              "right": "e"}.get(horizontal, "w")
                    justify = {"left": "left", "center": "center",
                               "right": "right"}.get(horizontal, "left")
                    wrap = bool(align and align.wrap_text)

                    widget = CellWidget(
                        self, coord, r, c, grid_row_of[r], text, w_px, h_px,
                        rs, cs, self._font(bold, min(size, 16)),
                        bg, fg, anchor, justify, wrap)
                    self.cells[coord] = widget
                    self._row_items.setdefault(widget.grid_row, []).append(widget.frame)
                    if coord in keep:
                        widget.set_new_text(keep[coord])

            self._compute_anchors()
            self._build_nav()
            self._apply_page(self.page)
            return len(self.cells)
        finally:
            wb.close()

    # -- Seiten -----------------------------------------------------------

    def _build_nav(self):
        for w in self.jump_frame.winfo_children():
            w.destroy()
        for n in range(len(self.pages)):
            ttk.Button(self.jump_frame, text=f"Seite {n + 1}", width=8,
                       command=lambda n=n: self.show_page(n)).pack(side="left", padx=2)

    def page_of(self, coord: str) -> int:
        cell = self.cells.get(coord)
        if cell:
            for n, rows in enumerate(self.pages):
                if cell.grid_row in rows:
                    return n
        return 0

    def show_page(self, n: int):
        n = max(0, min(n, len(self.pages) - 1))
        self._apply_page(n)

    def _apply_page(self, n: int):
        """Blendet alle Zeilen außer denen der Seite n aus."""
        self.page = n
        visible = set(self.pages[n])
        for idx, items in self._row_items.items():
            on = idx in visible
            g_min, l_min = self._row_min.get(idx, (0, 0))
            self.grid_frame.grid_rowconfigure(idx, minsize=g_min if on else 0)
            self.left_frame.grid_rowconfigure(idx, minsize=l_min if on else 0)
            for w in items:
                if isinstance(w, tk.Canvas):
                    w.grid_remove()  # Umbruchlinie: Seiten sind getrennt
                elif on:
                    w.grid()
                else:
                    w.grid_remove()
        total = len(self.pages)
        self.page_label.configure(text=f"Seite {n + 1} von {total}")
        self.prev_btn.configure(state="normal" if n > 0 else "disabled")
        self.next_btn.configure(state="normal" if n < total - 1 else "disabled")
        self.grid_frame.update_idletasks()
        self._update_scrollregions()
        self.canvas.xview_moveto(0)
        self.canvas.yview_moveto(0)

    def _row_label(self, text: str, height: int, fg: str, font, idx: int):
        """Zeilenkopf mit fester Höhe (wächst nicht mit der Schrift)."""
        box = tk.Frame(self.left_frame, width=ROW_HEADER_W, height=height,
                       background=HEADER_BG, highlightthickness=1,
                       highlightbackground=GRID_COLOR)
        box.pack_propagate(False)
        tk.Label(box, text=text, background=HEADER_BG, foreground=fg,
                 font=font, borderwidth=0).pack(fill="both", expand=True)
        box.grid(row=idx, column=0, sticky="nsew")
        return box

    def _draw_sep(self, canvas: tk.Canvas):
        canvas.delete("all")
        canvas.create_line(2, 4, canvas.winfo_width(), 4,
                           fill="#7A86A0", dash=(6, 4))

    def _compute_anchors(self):
        """Fuer jede Zelle die naechste Beschriftung (links, sonst oberhalb)."""
        texts = {c: w.orig_text for c, w in self.cells.items()
                 if not c.startswith(HF_PREFIX) and w.orig_text}
        for coord, w in self.cells.items():
            if coord.startswith(HF_PREFIX):
                continue
            anchor = ("", "")
            for cc in range(w.col - 1, 0, -1):
                t = texts.get(f"{idx_to_col(cc)}{w.row}")
                if t:
                    anchor = (t, f"{idx_to_col(cc)}{w.row}")
                    break
            if not anchor[0]:
                for rr in range(w.row - 1, 0, -1):
                    t = texts.get(f"{idx_to_col(w.col)}{rr}")
                    if t:
                        anchor = (t, f"{idx_to_col(w.col)}{rr}")
                        break
            self.anchors[coord] = anchor

    @staticmethod
    def _format_value(value) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def _update_scrollregions(self):
        for canvas in (self.canvas, self.top_canvas, self.left_canvas):
            canvas.configure(scrollregion=canvas.bbox("all"))

    # -- Interaktion ----------------------------------------------------

    def select(self, coord: Optional[str]):
        if coord is not None and coord not in self.cells:
            return
        previous = self.selected
        self.selected = coord
        if previous in self.cells:
            self.cells[previous]._refresh()
        if coord in self.cells:
            self.cells[coord]._refresh()
        self.on_select(coord)

    def set_new_value(self, coord: str, text: Optional[str]):
        if coord in self.cells:
            self.cells[coord].set_new_text(text)

    def orig_value(self, coord: str) -> str:
        cell = self.cells.get(coord)
        return cell.orig_text if cell else ""

    def anchor_of(self, coord: str) -> tuple[str, str]:
        return self.anchors.get(coord, ("", ""))

    def changed(self) -> dict[str, str]:
        """Alle geänderten Felder: Koordinate -> neuer Wert ('' = leeren)."""
        return {c: w.new_text for c, w in self.cells.items()
                if w.new_text is not None}

    def clear_all(self):
        for w in self.cells.values():
            w.set_new_text(None)

    # -- Suche ------------------------------------------------------------

    def search(self, text: str) -> int:
        for w in self.cells.values():
            w.flash_match(False)
        self._matches = []
        self._match_idx = -1
        if text.strip():
            needle = text.strip().lower()
            self._matches = [c for c, w in self.cells.items()
                             if needle in w.orig_text.lower() or needle in c.lower()]
        return len(self._matches)

    def next_match(self) -> Optional[str]:
        if not self._matches:
            return None
        if 0 <= self._match_idx < len(self._matches):
            self.cells[self._matches[self._match_idx]].flash_match(False)
        self._match_idx = (self._match_idx + 1) % len(self._matches)
        coord = self._matches[self._match_idx]
        self.cells[coord].flash_match(True)
        if self.page_of(coord) != self.page:
            self.show_page(self.page_of(coord))
        self.scroll_to(coord)
        self.select(coord)
        return coord

    def scroll_to(self, coord: str):
        cell = self.cells.get(coord)
        if not cell:
            return
        if self.page_of(coord) != self.page:
            self.show_page(self.page_of(coord))
            self.update_idletasks()
        total_h = max(self.grid_frame.winfo_height(), 1)
        total_w = max(self.grid_frame.winfo_width(), 1)
        y = max(0.0, (cell.frame.winfo_y() - 60) / total_h)
        x = max(0.0, (cell.frame.winfo_x() - 80) / total_w)
        self.canvas.yview_moveto(min(y, 1.0))
        self.canvas.xview_moveto(min(x, 1.0))

    # -- Scrollen ----------------------------------------------------------

    def _on_wheel(self, event):
        if getattr(event, "num", None) == 4 or getattr(event, "delta", 0) > 0:
            step = -3
        else:
            step = 3
        if getattr(event, "state", 0) & 0x1:  # Shift = horizontal
            self.canvas.xview_scroll(step, "units")
        else:
            self.canvas.yview_scroll(step, "units")
        return "break"
