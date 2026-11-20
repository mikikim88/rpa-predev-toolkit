"""Shared python-pptx helpers for RPA deliverables (PDD / SDD).

All geometry is in inches on a 16:9 (13.333 x 7.5) slide.
Any text containing "확인 필요" is automatically highlighted in red so that
missing information stands out to reviewers.
"""
from __future__ import annotations

import io
import math
import sys
from pathlib import Path
from typing import Any, Iterable, Sequence

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent))
from _template import template_bytes  # noqa: E402

FONT = "Malgun Gothic"
TBD = "확인 필요"

# Charcoal Minimal palette
C = {
    "primary": "36454F",
    "light": "F2F2F2",
    "dark": "212121",
    "highlight": "E63946",
    "success": "2A9D8F",
    "warn": "E9A23B",
    "muted": "7A858C",
    "white": "FFFFFF",
    "border": "D0D5D9",
    "zebra": "F7F8F9",
}

SLIDE_W = 13.333
SLIDE_H = 7.5
CONTENT_TOP = 1.5
CONTENT_BOTTOM = 6.85
CONTENT_H = CONTENT_BOTTOM - CONTENT_TOP
MARGIN_X = 0.6
CONTENT_W = SLIDE_W - 2 * MARGIN_X

LEFT, CENTER, RIGHT = PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT
TOP, MIDDLE, BOTTOM = MSO_ANCHOR.TOP, MSO_ANCHOR.MIDDLE, MSO_ANCHOR.BOTTOM


# --------------------------------------------------------------------------- #
# value helpers
# --------------------------------------------------------------------------- #
def blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and not v.strip()) or (isinstance(v, (list, dict)) and not v)


def fmt(v: Any, default: str = TBD) -> str:
    """Render any JSON value as display text. Empty -> '확인 필요'."""
    if blank(v):
        return default
    if isinstance(v, bool):
        return "예" if v else "아니오"
    if isinstance(v, float):
        return f"{v:,.1f}".rstrip("0").rstrip(".")
    if isinstance(v, int):
        return f"{v:,}"
    if isinstance(v, (list, tuple)):
        return "\n".join(fmt(x, default) for x in v)
    if isinstance(v, dict):
        return ", ".join(f"{k}: {fmt(x, default)}" for k, x in v.items())
    return str(v)


def num(v: Any) -> float | None:
    try:
        if isinstance(v, str):
            v = v.replace(",", "").strip()
        return float(v)
    except (TypeError, ValueError):
        return None


def is_tbd(text: str) -> bool:
    return TBD in str(text)


def rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h)


def grade_color(grade: str) -> str:
    return {"상": C["success"], "중": C["warn"], "하": C["highlight"]}.get(str(grade).strip(), C["muted"])


def impact_color(level: str) -> str:
    # for risk impact / automation difficulty: 상 = bad
    return {"상": C["highlight"], "중": C["warn"], "하": C["success"]}.get(str(level).strip(), C["muted"])


# --------------------------------------------------------------------------- #
# text measurement (rough, used for table pagination)
# --------------------------------------------------------------------------- #
def text_width_in(s: str, size: float) -> float:
    w = 0.0
    for ch in s:
        o = ord(ch)
        if o >= 0x2E80:
            w += 1.0
        elif ch == " ":
            w += 0.3
        elif ch.isupper() or ch.isdigit():
            w += 0.62
        else:
            w += 0.52
    return w * size / 72


def wrap_words(text: str, width_in: float, size: float, bold: bool = False) -> list[str]:
    """Greedy word wrap. PowerPoint breaks Hangul between any two characters
    (e.g. "등/록"), so we pre-break lines on spaces using a slightly
    pessimistic width estimate; PowerPoint then never needs to break a word."""
    factor = 1.08 if bold else 1.02
    words = str(text).split(" ")
    lines, cur = [], ""
    for w in words:
        cand = f"{cur} {w}" if cur else w
        if cur and text_width_in(cand, size) * factor > width_in:
            lines.append(cur)
            cur = w
        else:
            cur = cand
    lines.append(cur)
    return lines


def est_lines(text: str, width_in: float, size: float) -> int:
    usable = max(width_in - 0.18, 0.3)
    return sum(len(wrap_words(p, usable, size)) for p in str(text).split("\n"))


def est_row_h(cells: Sequence[str], widths: Sequence[float], size: float) -> float:
    lines = max(est_lines(c, w, size) for c, w in zip(cells, widths))
    return lines * size * 1.32 / 72 + 0.14


def paginate(rows: list[list[str]], widths: Sequence[float], size: float,
             max_h: float, header_h: float = 0.42) -> list[list[list[str]]]:
    pages, cur, h = [], [], header_h
    for r in rows:
        rh = est_row_h(r, widths, size)
        if cur and h + rh > max_h:
            pages.append(cur)
            cur, h = [], header_h
        cur.append(r)
        h += rh
    if cur or not pages:
        pages.append(cur)
    return pages


# --------------------------------------------------------------------------- #
# low level drawing
# --------------------------------------------------------------------------- #
def style_run(run, size: float = 12, bold: bool = False, color: str = C["dark"], italic: bool = False) -> None:
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    f.italic = italic
    f.color.rgb = rgb(color)
    f.name = FONT
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", FONT)


def _fill_frame(tf, content, size, bold, color, align, italic=False, bullets=False,
                auto_tbd=True, line_spacing=None, wrap_w: float | None = None):
    paras = content if isinstance(content, (list, tuple)) else str(content).split("\n")
    first = True
    for p_text in paras:
        opts: dict = {}
        if isinstance(p_text, tuple):
            p_text, opts = p_text
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = opts.get("align", align)
        if line_spacing:
            p.line_spacing = line_spacing
        if opts.get("space_before"):
            p.space_before = Pt(opts["space_before"])
        text = str(p_text)
        if bullets and text and not opts.get("no_bullet"):
            text = "• " + text
        c = opts.get("color", color)
        if auto_tbd and is_tbd(text):
            c = C["highlight"]
        r_size, r_bold = opts.get("size", size), opts.get("bold", bold)
        lines = wrap_words(text, wrap_w, r_size, r_bold) if wrap_w else [text]
        for li, line in enumerate(lines):
            if li:
                p.add_line_break()
            run = p.add_run()
            run.text = line
            style_run(run, r_size, r_bold, c, opts.get("italic", italic))


def add_rect(slide, x, y, w, h, fill: str | None = None, line: str | None = None,
             shape=MSO_SHAPE.RECTANGLE, line_w: float = 0.75):
    shp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(fill)
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = rgb(line)
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def add_text(slide, x, y, w, h, content, size: float = 12, bold: bool = False,
             color: str = C["dark"], align=LEFT, anchor=TOP, fill: str | None = None,
             line: str | None = None, italic: bool = False, bullets: bool = False,
             auto_tbd: bool = True, margin: float = 0.06, shape=MSO_SHAPE.RECTANGLE,
             line_spacing=None):
    shp = add_rect(slide, x, y, w, h, fill, line, shape)
    tf = shp.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Inches(margin))
    wrap_w = w - 2 * margin - (0.5 if shape != MSO_SHAPE.RECTANGLE else 0.02)
    _fill_frame(tf, content, size, bold, color, align, italic, bullets, auto_tbd, line_spacing, wrap_w)
    return shp


def add_table(slide, x, y, widths: Sequence[float], headers: Sequence[str], rows: list[list[str]],
              size: float = 10.5, header_fill: str = C["primary"], zebra: bool = True,
              row_colors: list[str | None] | None = None, cell_fills: dict | None = None,
              align: Sequence | None = None, bold_cols: Iterable[int] = ()):
    """Draw a table. row_colors: text colour per row. cell_fills: {(r, c): hex} (r is body index)."""
    bold_cols = set(bold_cols)
    n_rows = len(rows) + 1
    total_h = 0.42 + sum(est_row_h(r, widths, size) for r in rows)
    gshape = slide.shapes.add_table(n_rows, len(headers), Inches(x), Inches(y),
                                    Inches(sum(widths)), Inches(total_h))
    tbl = gshape.table
    # strip default table style banding
    tbl.first_row = True
    tbl.horz_banding = False
    for i, w in enumerate(widths):
        tbl.columns[i].width = Inches(w)
    tbl.rows[0].height = Inches(0.42)
    for r_i, r in enumerate(rows, start=1):
        tbl.rows[r_i].height = Inches(est_row_h(r, widths, size))

    def _cell(cell, text, fill, color, bold, al, width):
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(fill)
        cell.margin_left = cell.margin_right = Inches(0.07)
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MIDDLE
        tf = cell.text_frame
        tf.word_wrap = True
        _fill_frame(tf, text, size, bold, color, al, wrap_w=width - 0.16)

    for c_i, h in enumerate(headers):
        _cell(tbl.cell(0, c_i), h, header_fill, C["white"], True, CENTER, widths[c_i])
    for r_i, r in enumerate(rows):
        base = C["zebra"] if (zebra and r_i % 2 == 1) else C["white"]
        color = (row_colors[r_i] if row_colors and row_colors[r_i] else C["dark"])
        for c_i, val in enumerate(r):
            fill = (cell_fills or {}).get((r_i, c_i), base)
            al = align[c_i] if align else LEFT
            txt_color = C["white"] if (cell_fills and (r_i, c_i) in cell_fills) else color
            _cell(tbl.cell(r_i + 1, c_i), val, fill, txt_color, c_i in bold_cols, al, widths[c_i])
    return gshape


def add_card(slide, x, y, w, h, title: str, body, accent: str = C["primary"],
             title_size: float = 14, body_size: float = 11.5, bullets: bool = False):
    add_rect(slide, x, y, w, h, C["light"])
    add_rect(slide, x, y, 0.08, h, accent)
    add_text(slide, x + 0.2, y + 0.1, w - 0.3, 0.45, title, size=title_size, bold=True,
             color=accent if accent != C["light"] else C["dark"], anchor=MIDDLE)
    add_text(slide, x + 0.2, y + 0.6, w - 0.3, h - 0.7, body, size=body_size, bullets=bullets)


def add_badge(slide, x, y, w, h, text: str, fill: str, size: float = 11, color: str = C["white"]):
    return add_text(slide, x, y, w, h, text, size=size, bold=True, color=color, align=CENTER,
                    anchor=MIDDLE, fill=fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, auto_tbd=False)


def add_chevrons(slide, x, y, w, h, labels: Sequence[str], fills: Sequence[str] | None = None,
                 size: float = 12, gap: float = 0.06):
    n = max(len(labels), 1)
    cw = (w - gap * (n - 1)) / n
    for i, lab in enumerate(labels):
        shape = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
        fill = fills[i] if fills else C["primary"]
        add_text(slide, x + i * (cw + gap), y, cw + (0.15 if i else 0), h, lab, size=size, bold=True,
                 color=C["white"], align=CENTER, anchor=MIDDLE, fill=fill, shape=shape,
                 auto_tbd=False, margin=0.25 if i else 0.1)


def add_arrow_down(slide, x, y, w=0.4, h=0.35, fill=C["muted"]):
    return add_rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.DOWN_ARROW)


def add_score_bars(slide, x, y, w, items: Sequence[tuple[str, float | None, float]],
                   row_h: float = 0.55, label_w: float = 1.9):
    """items: (label, score, max). Bar colour by score ratio."""
    bar_w = w - label_w - 0.8
    for i, (label, score, mx) in enumerate(items):
        yy = y + i * row_h
        add_text(slide, x, yy, label_w, row_h - 0.1, label, size=11.5, bold=True, anchor=MIDDLE)
        add_rect(slide, x + label_w, yy + 0.12, bar_w, row_h - 0.34, C["light"])
        if score is not None:
            ratio = max(0.0, min(1.0, score / mx))
            col = C["success"] if ratio >= 0.8 else C["warn"] if ratio >= 0.5 else C["highlight"]
            add_rect(slide, x + label_w, yy + 0.12, max(bar_w * ratio, 0.05), row_h - 0.34, col)
            txt, col = f"{score:g} / {mx:g}", C["dark"]
        else:
            txt, col = "미평가", C["highlight"]
        add_text(slide, x + label_w + bar_w + 0.05, yy, 0.8, row_h - 0.1, txt, size=11, bold=True,
                 color=col, anchor=MIDDLE)


# --------------------------------------------------------------------------- #
# deck
# --------------------------------------------------------------------------- #
class Deck:
    def __init__(self, doc_label: str, task_name: str):
        # embedded template: immune to DRM-encrypted python-pptx installs
        self.prs = Presentation(io.BytesIO(template_bytes()))
        self.prs.slide_width = Inches(SLIDE_W)
        self.prs.slide_height = Inches(SLIDE_H)
        self.doc_label = doc_label
        self.task_name = task_name
        self.n = 0
        self._layout = self.prs.slide_layouts[6]

    def blank(self):
        self.n += 1
        return self.prs.slides.add_slide(self._layout)

    def footer(self, s):
        add_rect(s, MARGIN_X, 7.0, CONTENT_W, 0.012, C["border"])
        add_text(s, MARGIN_X, 7.05, 9, 0.3, f"{self.doc_label}  |  {self.task_name}", size=9,
                 color=C["muted"], auto_tbd=False)
        add_text(s, SLIDE_W - MARGIN_X - 1.5, 7.05, 1.5, 0.3, str(self.n), size=9, color=C["muted"],
                 align=RIGHT, auto_tbd=False)

    def slide(self, title: str, caption: str | None = None):
        s = self.blank()
        add_rect(s, 0, 0, SLIDE_W, 0.1, C["primary"])
        add_rect(s, MARGIN_X, 0.48, 0.09, 0.55, C["highlight"])
        add_text(s, MARGIN_X + 0.22, 0.4, 11.5, 0.7, title, size=24, bold=True, color=C["primary"],
                 anchor=MIDDLE, auto_tbd=False)
        if caption:
            add_text(s, MARGIN_X + 0.22, 1.02, 12, 0.36, caption, size=11, color=C["muted"],
                     italic=True)
        self.footer(s)
        return s

    def cover(self, title: str, subtitle: str, meta: list[tuple[str, str]], tag: str):
        s = self.blank()
        split = 8.4
        add_rect(s, split, 0, SLIDE_W - split, SLIDE_H, C["primary"])
        add_rect(s, split, 0, 0.12, SLIDE_H, C["highlight"])
        add_text(s, split + 0.5, 2.3, SLIDE_W - split - 0.9, 1.6, tag, size=60, bold=True,
                 color=C["white"], align=LEFT, anchor=MIDDLE, auto_tbd=False)
        add_text(s, split + 0.55, 3.8, SLIDE_W - split - 0.9, 0.5, self.doc_label, size=14,
                 color="CBD3D8", auto_tbd=False)
        add_text(s, 0.8, 1.5, 7.2, 0.6, title, size=20, bold=False, color=C["muted"], auto_tbd=False)
        add_text(s, 0.8, 2.1, 7.3, 1.9, subtitle, size=34, bold=True, color=C["primary"],
                 anchor=TOP)
        add_rect(s, 0.8, 4.25, 1.2, 0.05, C["highlight"])
        rows = [(k, fmt(v)) for k, v in meta]
        for i, (k, v) in enumerate(rows):
            add_text(s, 0.8, 4.55 + i * 0.4, 1.6, 0.36, k, size=12, bold=True, color=C["muted"],
                     anchor=MIDDLE)
            add_text(s, 2.4, 4.55 + i * 0.4, 5.6, 0.36, v, size=12, anchor=MIDDLE)
        return s

    def table_slides(self, title: str, headers, rows, widths, caption: str | None = None,
                     size: float = 10.5, top: float = CONTENT_TOP, max_h: float | None = None,
                     row_colors=None, cell_fill_fn=None, align=None, bold_cols=(), before=None,
                     zebra: bool = True):
        """Paginate a table over as many slides as needed. Returns list of slides.

        cell_fill_fn(row) -> {col_index: hex} for coloured cells.
        before(slide, page_index) lets callers draw extras (legend etc.)."""
        max_h = max_h or (CONTENT_BOTTOM - top)
        if not rows:
            rows = [[TBD] + [""] * (len(headers) - 1)]
        pages = paginate(rows, widths, size, max_h)
        slides, offset = [], 0
        for p_i, page in enumerate(pages):
            t = title if len(pages) == 1 else f"{title} ({p_i + 1}/{len(pages)})"
            s = self.slide(t, caption)
            if before:
                before(s, p_i)
            rc = row_colors[offset:offset + len(page)] if row_colors else None
            fills = {}
            if cell_fill_fn:
                for r_i, r in enumerate(page):
                    for c_i, col in (cell_fill_fn(offset + r_i) or {}).items():
                        fills[(r_i, c_i)] = col
            add_table(s, MARGIN_X, top, widths, headers, page, size=size, row_colors=rc,
                      cell_fills=fills, align=align, bold_cols=bold_cols, zebra=zebra)
            offset += len(page)
            slides.append(s)
        return slides

    def save(self, path: str) -> None:
        self.prs.save(path)
