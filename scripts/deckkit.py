"""Shared helpers for the PowerPoint decks (language aware)."""
import os, sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
sys.path.insert(0, os.path.dirname(__file__))
from strings import tr

DARK, MID, LIGHT, GOLD, CREAM, WHITE, GREY = (RGBColor.from_string(c) for c in
    ("0F3D2E", "2E7D5B", "E8F3EC", "D4A017", "FBF7EE", "FFFFFF", "5F6B66"))
# Nirmala UI ships with Windows/Office and covers Latin, Devanagari and Gujarati.
FONT = {"en": "Arial", "hi": "Nirmala UI", "gu": "Nirmala UI"}
SUFFIX = {"en": "", "hi": "_Hindi", "gu": "_Gujarati"}
CONTACT_NAME = "Harman Ram"
CONTACT_PHONE = "+91 7777999896"

class Deck:
    def __init__(self, lang):
        self.lang = lang
        self.A = f"assets/{lang}/"
        self.prs = Presentation(); self.prs.slide_width = Inches(13.333); self.prs.slide_height = Inches(7.5)
        self.BL = self.prs.slide_layouts[6]
    def T(self, s): return tr(s, self.lang)
    def rs(self, s):
        """Currency: 'Rs 20,000' in English, '₹20,000' in Hindi/Gujarati."""
        return s if self.lang == "en" else s.replace("Rs ", "₹")
    def slide(self): return self.prs.slides.add_slide(self.BL)
    def font(self, f, size, color, bold):
        f.name = FONT[self.lang]; f.size = Pt(size); f.color.rgb = color; f.bold = bold
        rPr = f._rPr
        for tag in ("a:cs", "a:ea"):
            el = rPr.find(qn(tag))
            if el is None: el = rPr.makeelement(qn(tag), {}); rPr.append(el)
            el.set("typeface", FONT[self.lang])
    def rect(self, s, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
        r = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        r.fill.solid(); r.fill.fore_color.rgb = fill; r.line.fill.background(); r.shadow.inherit = False
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE: r.adjustments[0] = 0.07
        return r
    def text(self, s, x, y, w, h, t, size=18, color=DARK, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
        tf.word_wrap = True; tf.vertical_anchor = anchor; tf.margin_left = tf.margin_right = Inches(0.05)
        for i, ln in enumerate(t if isinstance(t, list) else [t]):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            ln, sz, col, b = ln if isinstance(ln, tuple) else (ln, size, color, bold)
            p.alignment = align; p.space_after = Pt(6)
            r = p.add_run(); r.text = ln; self.font(r.font, sz, col, b)
        return tb
    def bg(self, s, c=CREAM): self.rect(s, 0, 0, 13.333, 7.5, c)
    def overlay(self, s, alpha="74000"):
        s.shapes.add_picture(self.A + "site_layout.png", 0, 0, self.prs.slide_width, self.prs.slide_height)
        ov = self.rect(s, 0, 0, 13.333, 7.5, DARK); clr = ov.fill._xPr.find(qn("a:solidFill"))[0]
        clr.append(clr.makeelement(qn("a:alpha"), {"val": alpha}))
    def heading(self, s, t, sub=None, dark=False):
        self.rect(s, 0.6, 0.55, 0.12, 0.75, GOLD)
        self.text(s, 0.85, 0.42, 11.8, 0.7, t, 34, WHITE if dark else DARK, True)
        if sub: self.text(s, 0.87, 1.08, 11.8, 0.5, sub, 16, LIGHT if dark else GREY)
    def footer(self, s, n, label, dark=False):
        c = LIGHT if dark else GREY
        self.text(s, 0.6, 7.0, 9, 0.35, label, 10, c)
        self.text(s, 12.0, 7.0, 0.8, 0.35, str(n), 10, c, align=PP_ALIGN.RIGHT)
    def card(self, s, x, y, w, h, big, small, fill=WHITE, bigc=DARK, smallc=GREY, bs=30):
        self.rect(s, x, y, w, h, fill, MSO_SHAPE.ROUNDED_RECTANGLE)
        self.text(s, x + 0.15, y + 0.15, w - 0.3, h * 0.5, big, bs, bigc, True, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        self.text(s, x + 0.15, y + h * 0.56, w - 0.3, h * 0.42, small, 14, smallc, False, PP_ALIGN.CENTER)
    def contact(self, s, x, y, w, h, place):
        self.rect(s, x, y, w, h, WHITE, MSO_SHAPE.ROUNDED_RECTANGLE)
        self.text(s, x, y, w, h, [(self.T("Contact: {name}").format(name=self.T(CONTACT_NAME)), 24, DARK, True),
                                  (self.T("Call: {phone}").format(phone=CONTACT_PHONE), 24, MID, True),
                                  (self.T(place), 15, GREY, False)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    def table(self, s, rows, x, y, w, h, colw=None, size=17, highlight=lambda r: False, hl=None, left_first=False, boldcol=None):
        tb = s.shapes.add_table(len(rows), len(rows[0]), Inches(x), Inches(y), Inches(w), Inches(h)).table
        if colw:
            for i, cw in enumerate(colw): tb.columns[i].width = Inches(cw)
        for r, row in enumerate(rows):
            for c, v in enumerate(row):
                cell = tb.cell(r, c); cell.text = v; cell.fill.solid(); cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.fill.fore_color.rgb = DARK if r == 0 else (hl if highlight(r) else WHITE)
                p = cell.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if (left_first and c == 0) else PP_ALIGN.CENTER
                b = r == 0 or highlight(r) or (left_first and c == 0) or (boldcol is not None and c == boldcol)
                if p.runs: self.font(p.runs[0].font, size, WHITE if r == 0 else DARK, b)
        return tb
    def save(self, base):
        out = f"output/{base}{SUFFIX[self.lang]}.pptx"; self.prs.save(out); print("saved", out); return out
