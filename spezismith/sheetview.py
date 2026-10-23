"""
SpeziSmith - SheetView

Maßstabsgetreue, klickbare Nachbildung der Spezifikations-Vorlage:
Spaltenbreiten, Zeilenhöhen, verbundene Zellen, Schrift- und Füllfarben
werden aus der Excel-Datei übernommen. Ein Klick auf ein Feld wählt es aus,
eingetragene Änderungen werden direkt im Layout sichtbar (WYSIWYG).
"""

from __future__ import annotations

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
from typing import Callable, Optional

import openpyxl

from .engine import idx_to_col

ORANGE = "#F08C21"
NAVY = "#1F2942"
CHANGED_BG = "#FFE3C2"
MATCH_BG = "#FFF59D"
HEADER_BG = "#E9EDF3"
GRID_COLOR = "#B9BFC9"

ROW_HEADER_W = 44
COL_HEADER_H = 22
DEFAULT_COL_WIDTH = 8.43   # Excel-Standardbreite
DEFAULT_ROW_HEIGHT = 16.5  # Punkte


def _px_width(excel_width: float) -> int:
    return max(10, int(excel_width * 7 + 6))


def _px_height(points: float) -> int:
    return max(14, int(points * 96 / 72))


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
    """Ein klickbares Feld im Tabellenraster."""

    def __init__(self, sheet: "SheetView", coord: str, row: int, col: int,
                 text: str, w_px: int, h_px: int, rowspan: int, colspan: int,
                 font: tkfont.Font, bg: str, fg: str, anchor: str,
                 justify: str, wrap: bool):
        self.sheet = sheet
        self.coord = coord
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

        self.frame.grid(row=row, column=col, rowspan=rowspan,
                        columnspan=colspan, sticky="nsew")
        for w in (self.frame, self.label):
            w.bind("<Button-1>", lambda _e, c=coord: sheet.select(c))
        sheet.register_tooltip(self.label, f"{coord}: {text or '(leer)'}")

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
    """Scrollbare, klickbare Vorlagen-Ansicht mit fixierten Kopfleisten."""

    def __init__(self, master, on_select: Callable[[Optional[str]], None]):
        super().__init__(master)
        self.on_select = on_select
        self.cells: dict[str, CellWidget] = {}
        self.selected: Optional[str] = None
        self._fonts: dict[tuple, tkfont.Font] = {}
        self._matches: list[str] = []
        self._match_idx = -1

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

        self.grid_frame = tk.Frame(self.canvas, background="white")
        self.top_frame = tk.Frame(self.top_canvas, background=HEADER_BG)
        self.left_frame = tk.Frame(self.left_canvas, background=HEADER_BG)
        self._grid_win = self.canvas.create_window((0, 0), window=self.grid_frame,
                                                   anchor="nw")
        self._top_win = self.top_canvas.create_window((0, 0), window=self.top_frame,
                                                      anchor="nw")
        self._left_win = self.left_canvas.create_window((0, 0), window=self.left_frame,
                                                        anchor="nw")
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

    def _font(self, bold: bool, size: float) -> tkfont.Font:
        key = (bold, round(size, 1))
        if key not in self._fonts:
            self._fonts[key] = tkfont.Font(
                family="Segoe UI", size=int(round(size)),
                weight="bold" if bold else "normal")
        return self._fonts[key]

    def load_template(self, path: str, keep: Optional[dict] = None):
        """Lädt eine Vorlage und baut das Raster komplett neu auf."""
        for cell in self.cells.values():
            cell.frame.destroy()
        for w in self.top_frame.winfo_children() + self.left_frame.winfo_children():
            w.destroy()
        self.cells.clear()
        self._tooltips.clear()
        self.selected = None
        self._matches = []
        keep = keep or {}

        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            ws = wb.worksheets[0]
            max_row, max_col = ws.max_row, ws.max_column

            # Spaltenbreiten / Zeilenhöhen in Pixeln
            col_px = [ROW_HEADER_W]
            for c in range(1, max_col + 1):
                dim = ws.column_dimensions.get(idx_to_col(c))
                width = dim.width if (dim and dim.width) else DEFAULT_COL_WIDTH
                col_px.append(_px_width(width))
            row_px = [COL_HEADER_H]
            for r in range(1, max_row + 1):
                dim = ws.row_dimensions.get(r)
                height = dim.height if (dim and dim.height) else DEFAULT_ROW_HEIGHT
                row_px.append(_px_height(height))

            # Verbundene Zellen
            span: dict[str, tuple[int, int]] = {}
            covered: set[str] = set()
            for rng in ws.merged_cells.ranges:
                anchor = f"{idx_to_col(rng.min_col)}{rng.min_row}"
                span[anchor] = (rng.max_row - rng.min_row + 1,
                                rng.max_col - rng.min_col + 1)
                for r in range(rng.min_row, rng.max_row + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        if (r, c) != (rng.min_row, rng.min_col):
                            covered.add(f"{idx_to_col(c)}{r}")

            # Kopfleisten
            for c in range(1, max_col + 1):
                tk.Label(self.top_frame, text=idx_to_col(c), background=HEADER_BG,
                         foreground="#4A5568", font=self._font(True, 9),
                         width=col_px[c] // 7, relief="solid", borderwidth=1,
                         ).grid(row=0, column=c, sticky="nsew",
                                ipady=2)
                self.top_frame.grid_columnconfigure(c, minsize=col_px[c])
            for r in range(1, max_row + 1):
                tk.Label(self.left_frame, text=str(r), background=HEADER_BG,
                         foreground="#4A5568", font=self._font(True, 9),
                         relief="solid", borderwidth=1,
                         ).grid(row=r, column=0, sticky="nsew")
                self.left_frame.grid_rowconfigure(r, minsize=row_px[r] + 1)

            # Raster-Konfiguration (Mindestgrößen)
            for c in range(1, max_col + 1):
                self.grid_frame.grid_columnconfigure(c, minsize=col_px[c])
            for r in range(1, max_row + 1):
                self.grid_frame.grid_rowconfigure(r, minsize=row_px[r])

            # Zellen
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    coord = f"{idx_to_col(c)}{r}"
                    if coord in covered:
                        continue
                    cell = ws.cell(row=r, column=c)
                    rs, cs = span.get(coord, (1, 1))
                    w_px = sum(col_px[c:c + cs]) + (cs - 1)
                    h_px = sum(row_px[r:r + rs]) + (rs - 1)

                    value = cell.value
                    text = self._format_value(value)

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
                    if align and align.vertical == "center":
                        anchor = {"w": "w", "center": "center", "e": "e"}[anchor]
                    wrap = bool(align and align.wrap_text)

                    widget = CellWidget(self, coord, r, c, text, w_px, h_px,
                                        rs, cs, self._font(bold, min(size, 16)),
                                        bg, fg, anchor, justify, wrap)
                    self.cells[coord] = widget

                    if coord in keep:
                        widget.set_new_text(keep[coord])

            self.grid_frame.update_idletasks()
            self._update_scrollregions()
            return len(self.cells)
        finally:
            wb.close()

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
        if self._match_idx >= 0 and self._match_idx < len(self._matches):
            self.cells[self._matches[self._match_idx]].flash_match(False)
        self._match_idx = (self._match_idx + 1) % len(self._matches)
        coord = self._matches[self._match_idx]
        self.cells[coord].flash_match(True)
        self.scroll_to(coord)
        self.select(coord)
        return coord

    def scroll_to(self, coord: str):
        cell = self.cells.get(coord)
        if not cell:
            return
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
