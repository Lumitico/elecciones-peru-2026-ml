#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Generador de la versión Word/DOCX camera-ready del Paper LACCEI 2026.

Fuente de verdad: Paper_LACCEI_2026.tex (versión LaTeX corregida y ya
compilada a PDF). Este script parsea el .tex y produce un .docx espejo:

- Página 1 (1 columna): línea ISBN/ISSN/DOI, título EN, autores, Abstract EN,
  Keywords EN.
- Página 2 en adelante (2 columnas): título ES + autores, Resumen ES,
  Palabras clave ES, cuerpo, ecuaciones OMML (1)-(5), tablas y figuras,
  Referencias 1-33 en orden de primera citación.
- Tablas anchas (table*) en secciones continuas de 1 columna a 7.25";
  tablas de columna a 3.5". El <w:cols> existente se MODIFICA (nunca se
  agrega uno segundo: ese era el bug de la versión anterior).
- Pie de página oficial del congreso + número de página en todas las páginas.

Salida: Paper_LACCEI_2026.docx
"""
import re
import os

import docx
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import qn, nsdecls

try:
    from PIL import Image
except Exception:
    Image = None

BASE = os.path.dirname(os.path.abspath(__file__))
TEX = os.path.join(BASE, 'Paper_LACCEI_2026.tex')
OUT = os.path.join(BASE, 'Paper_LACCEI_2026.docx')
OLD = os.path.join(BASE, 'compilar_paper_laccei_2026.py')

# ── Constantes de formato (template LACCEI 2026) ──────────────────────────────
MARGIN_TOP = Inches(0.75)
MARGIN_BOTTOM = Inches(1.0)
MARGIN_SIDE = Inches(0.625)
PAGE_W = Inches(8.5)
PAGE_H = Inches(11.0)
COLW_FULL = 7.25          # ancho de texto (8.5 - 2*0.625)
COLW_NARROW = 3.5         # columna de texto (7.25 - 0.25)/2
COL_GAP_TWIPS = 360       # 0.25"

FOOTER_TEXT = ('24th LACCEI International Multi-Conference for Engineering, '
               'Education, and Technology: \u201cEngineering without Borders: '
               'Artificial Intelligence, Knowledge, Innovation, and Alliances '
               'for a Future from the Americas\u201d. Hybrid Event, '
               'Santiago-Chile, July 15 - 17, 2026.')
DOI_TEXT = ('ISBN: 978-628-96613-3-0. ISSN: 2414-6390. Digital Object '
            'Identifier: https://dx.doi.org/10.18687/LACCEI2026.1.1.XXXX')

ROMAN = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
LETTERS = 'ABCDEFGHIJ'

NBSP = '\u00a0'
THIN = ' '  # espacio normal (U+2009 puede no existir en Times New Roman)

# Anchos de columna (pulgadas) por tabla, sumando 7.25" o 3.5"
TABLE_WIDTHS = {
    'tab:antecedentes': [1.15, 0.95, 1.00, 1.05, 1.10, 2.00],
    'tab:descriptivos': [1.50, 0.62, 0.62, 0.76],
    'tab:empates':      [1.05, 0.85, 0.95, 0.65],
    'tab:modelos':      [1.42, 1.42, 0.33, 0.88, 1.08, 1.08, 1.04],
    'tab:estabilidad':  [1.70, 0.40, 0.70, 0.70],
    'tab:perfiles':     [1.55, 1.35, 1.35, 1.35, 1.65],
}
TABLE_FONTS = {
    'tab:antecedentes': 7.0, 'tab:descriptivos': 7.8, 'tab:empates': 7.0,
    'tab:modelos': 7.0, 'tab:estabilidad': 7.8, 'tab:perfiles': 7.8,
}
SMALL_CAPS = hasattr(docx.text.run.Font, 'small_caps')

# ══════════════════════════════════════════════════════════════════════════════
# Utilidades XML
# ══════════════════════════════════════════════════════════════════════════════

def set_columns(section, num, space=None):
    """MODIFICA el <w:cols> existente del sectPr (nunca agrega uno segundo)."""
    sectPr = section._sectPr
    cols_list = sectPr.findall(qn('w:cols'))
    for extra in cols_list[1:]:
        sectPr.remove(extra)
    cols = cols_list[0] if cols_list else None
    if cols is None:
        cols = OxmlElement('w:cols')
        grid = sectPr.find(qn('w:docGrid'))
        if grid is not None:
            grid.addprevious(cols)
        else:
            sectPr.append(cols)
    cols.set(qn('w:num'), str(num))
    if space is not None:
        cols.set(qn('w:space'), str(space))
    else:
        cols.attrib.pop(qn('w:space'), None)
    return section


def add_section(doc, kind, num_cols):
    s = doc.add_section(kind)
    s.top_margin, s.bottom_margin = MARGIN_TOP, MARGIN_BOTTOM
    s.left_margin = s.right_margin = MARGIN_SIDE
    s.page_width, s.page_height = PAGE_W, PAGE_H
    set_columns(s, num_cols, COL_GAP_TWIPS if num_cols >= 2 else None)
    return s


def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:shd')):
        tcPr.remove(old)
    tcPr.append(parse_xml(
        '<w:shd %s w:val="clear" w:fill="%s"/>' % (nsdecls('w'), fill_hex)))


def set_cell_margins(cell, top=30, bottom=30, left=50, right=50):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcMar')):
        tcPr.remove(old)
    tcPr.append(parse_xml(
        '<w:tcMar %s>'
        '<w:top w:w="%d" w:type="dxa"/><w:bottom w:w="%d" w:type="dxa"/>'
        '<w:left w:w="%d" w:type="dxa"/><w:right w:w="%d" w:type="dxa"/>'
        '</w:tcMar>' % (nsdecls('w'), top, bottom, left, right)))


def set_row_props(row, header=False):
    trPr = row._tr.get_or_add_trPr()
    for tag in ('w:cantSplit', 'w:tblHeader'):
        for old in trPr.findall(qn(tag)):
            trPr.remove(old)
    trPr.append(parse_xml('<w:cantSplit %s/>' % nsdecls('w')))
    if header:
        trPr.append(parse_xml('<w:tblHeader %s/>' % nsdecls('w')))


def fix_table_layout(table, total_inches):
    tblPr = table._tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout'):
        for old in tblPr.findall(qn(tag)):
            tblPr.remove(old)
    dxa = int(total_inches * 1440)
    tblW = parse_xml('<w:tblW %s w:w="%d" w:type="dxa"/>'
                     % (nsdecls('w'), dxa))
    # orden de esquema: tblStyle, tblW, jc, ..., tblLayout, ..., tblLook
    style = tblPr.find(qn('w:tblStyle'))
    if style is not None:
        style.addnext(tblW)
    else:
        tblPr.insert(0, tblW)
    tblL = parse_xml('<w:tblLayout %s w:type="fixed"/>' % nsdecls('w'))
    anchor = None
    for tag in ('w:tblCellMar', 'w:tblLook'):
        e = tblPr.find(qn(tag))
        if e is not None:
            anchor = e
            break
    if anchor is not None:
        anchor.addprevious(tblL)
    else:
        tblPr.append(tblL)
    table.autofit = False


def add_runs(par, runs, size=None, font='Times New Roman'):
    for r in runs:
        if r.get('br'):
            par.add_run().add_break()
            continue
        run = par.add_run(r.get('t', ''))
        f = run.font
        f.name = 'Courier New' if r.get('mono') else font
        if r.get('mono'):
            f.size = Pt(9)
        elif size is not None:
            f.size = Pt(size)
        if r.get('b'):
            run.bold = True
        if r.get('i'):
            run.italic = True
        if r.get('sup'):
            f.superscript = True
        if r.get('sub'):
            f.subscript = True
        if r.get('sc') and SMALL_CAPS:
            try:
                f.small_caps = True
            except Exception:
                pass
    return par


# ══════════════════════════════════════════════════════════════════════════════
# Parser LaTeX (inline)
# ══════════════════════════════════════════════════════════════════════════════

def read_group(s, i):
    """s[i] debe ser '{'. Devuelve (contenido, índice tras la llave final)."""
    if s[i] != '{':
        raise ValueError('se esperaba { en idx %d: %r' % (i, s[i:i + 30]))
    depth = 0
    j = i
    n = len(s)
    while j < n:
        c = s[j]
        if c == '\\' and j + 1 < n:
            j += 2
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    raise ValueError('llaves desbalanceadas')


ACCENTS = {'a': 'á', 'e': 'é', 'i': 'í', 'o': 'ó', 'u': 'ú', 'n': 'ń',
           'A': 'Á', 'E': 'É', 'I': 'Í', 'O': 'Ó', 'U': 'Ú', 'N': 'Ń',
           'c': 'ç', 'C': 'Ç'}

GREEK = {'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'sigma': 'σ',
         'mu': 'μ', 'lambda': 'λ', 'pi': 'π', 'rho': 'ρ', 'tau': 'τ'}

SYMBOLS = {'ge': '≥', 'le': '≤', 'neq': '≠', 'times': '×', 'circ': '°',
           'uparrow': '↑', 'downarrow': '↓', 'pm': '±', 'approx': '≈',
           'rightarrow': '→', 'cdot': '·', '%': '%', '_': '_', '&': '&',
           '$': '$', '{': '{', '}': '}', '#': '#'}


class TexCtx:
    def __init__(self):
        self.cite_map = {}    # key -> número
        self.tab_map = {}     # label -> romano
        self.fig_map = {}     # label -> número
        self.eq_map = {}      # label -> número

    def cite(self, keys):
        nums = []
        for k in [k.strip() for k in keys.split(',') if k.strip()]:
            if k not in self.cite_map:
                raise KeyError('cita desconocida: %s' % k)
            nums.append(self.cite_map[k])
        return fmt_cites(nums)

    def ref(self, label):
        if label in self.tab_map:
            return self.tab_map[label]
        if label in self.fig_map:
            return str(self.fig_map[label])
        if label in self.eq_map:
            return str(self.eq_map[label])
        raise KeyError('ref desconocida: %s' % label)


def fmt_cites(nums):
    """[1], [2]; rangos de 3+ consecutivos como [1]–[3]."""
    out = []
    i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        span = j - i
        if span >= 2:
            out.append('%d\u2013%d' % (nums[i], nums[j]))
        elif span == 1:
            out.append(str(nums[i]))
            out.append(str(nums[j]))
        else:
            out.append(str(nums[i]))
        i = j + 1
    return '[' + '], ['.join(out) + ']'


def parse_math(m):
    """Matemática inline -> runs (variables en cursiva)."""
    runs = []

    def add(t, ital=False, sup=False, sub=False, b=False):
        if t:
            runs.append({'t': t, 'i': ital, 'sup': sup, 'sub': sub, 'b': b})

    i = 0
    n = len(m)
    while i < n:
        c = m[i]
        if c == '\\':
            mm = re.match(r'\\([a-zA-Z]+)|\\(.)', m[i:])
            if not mm:
                i += 1
                continue
            cmd = mm.group(1) or mm.group(2)
            i += mm.end()
            if cmd in ('mathrm', 'text', 'mathbf', 'mathit', 'mathcal'):
                g, i = read_group(m, i)
                sub_runs = parse_math(g)
                ital = cmd == 'mathit'
                bold = cmd in ('mathbf', 'mathcal')
                for r in sub_runs:
                    if not r.get('sub') and not r.get('sup'):
                        r['i'] = ital
                    r['b'] = bold or r.get('b')
                runs.extend(sub_runs)
            elif cmd in GREEK:
                add(GREEK[cmd], ital=True)
            elif cmd in SYMBOLS:
                add(SYMBOLS[cmd])
            elif cmd == ',':
                add(THIN)
            elif cmd in (';', ' ', '!'):
                add(' ')
            elif cmd in ('sum', 'prod'):
                add('∑')
            else:
                pass  # comando desconocido en math: se descarta
        elif c in ('^', '_'):
            sup = c == '^'
            i += 1
            if i < n and m[i] == '{':
                g, i = read_group(m, i)
            else:
                g = m[i]
                i += 1
            sub_runs = parse_math(g)
            for r in sub_runs:
                r.pop('sup', None)
                r.pop('sub', None)
                r['sup' if sup else 'sub'] = True
            runs.extend(sub_runs)
        elif c.isalpha():
            mm = re.match(r'[a-zA-Z]+', m[i:])
            add(mm.group(0), ital=True)
            i += mm.end()
        elif c.isspace():
            add(' ')
            i += 1
        else:
            add(c)
            i += 1
    return runs


def parse_inline(s, ctx):
    """Parser inline LaTeX -> lista de runs."""
    runs = []

    def add(**kw):
        runs.append(kw)

    buf = []

    def flush():
        if buf:
            add(t=''.join(buf))
            del buf[:]

    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c == '\\':
            mm = re.match(r'\\([a-zA-Z]+)\*?|\\(.)', s[i:])
            if not mm:
                i += 1
                continue
            cmd = mm.group(1) or mm.group(2)
            i += mm.end()
            flush()
            if cmd in ('textbf', 'textit', 'emph', 'texttt', 'textsc',
                       'textnormal', 'text', 'mathrm', 'mathbf', 'mathit',
                       'textsuperscript'):
                g, i = read_group(s, i)
                sub = parse_inline(g, ctx)
                if cmd == 'textbf':
                    for r in sub:
                        r['b'] = True
                elif cmd in ('textit', 'emph', 'mathit'):
                    for r in sub:
                        r['i'] = True
                elif cmd == 'texttt':
                    for r in sub:
                        r['mono'] = True
                elif cmd == 'textsuperscript':
                    for r in sub:
                        r['sup'] = True
                elif cmd == 'textsc':
                    for r in sub:
                        r['sc'] = True
                runs.extend(sub)
            elif cmd == 'href':
                _, i = read_group(s, i)          # URL: se descarta
                g2, i = read_group(s, i)
                runs.extend(parse_inline(g2, ctx))
            elif cmd == 'url':
                g, i = read_group(s, i)
                add(t=g)
            elif cmd == 'cite':
                g, i = read_group(s, i)
                add(t=ctx.cite(g))
            elif cmd == 'ref':
                g, i = read_group(s, i)
                add(t=ctx.ref(g))
            elif cmd == "'":
                if i < n:
                    ch = s[i]
                    if ch == '{':
                        g, i = read_group(s, i)
                        ch = g
                    else:
                        i += 1
                    base = ch[0] if ch else ''
                    buf.append(ACCENTS.get(base, base + '\u0301'))
                    if len(ch) > 1:
                        buf.append(ch[1:])
            elif cmd in SYMBOLS:
                buf.append(SYMBOLS[cmd])
            elif cmd in (',', ';', ':', ' '):
                buf.append(THIN if cmd == ',' else ' ')
            elif cmd in ('quad', 'qquad'):
                buf.append('\u2003' * (1 if cmd == 'quad' else 2))
            elif cmd == '~':
                buf.append(NBSP)
            elif cmd == '\\':
                add(br=True)
            else:
                # Comando desconocido: consumir grupos {} que le sigan y descartar
                while i < n and s[i] == '{':
                    _, i = read_group(s, i)
        elif c == '$':
            flush()
            j = s.find('$', i + 1)
            if j == -1:
                buf.append(c)
                i += 1
            else:
                runs.extend(parse_math(s[i + 1:j]))
                i = j + 1
        elif s.startswith('``', i):
            flush()
            buf.append('\u201c')
            i += 2
        elif s.startswith("''", i):
            flush()
            buf.append('\u201d')
            i += 2
        elif s.startswith('---', i):
            flush()
            buf.append('\u2014')
            i += 3
        elif s.startswith('--', i):
            flush()
            buf.append('\u2013')
            i += 2
        elif c == '~':
            buf.append(NBSP)
            i += 1
        elif c == '%':
            j = s.find('\n', i)
            i = n if j == -1 else j + 1
        elif c in ('{', '}'):
            i += 1  # agrupación vacía: se descarta
        else:
            buf.append(c)
            i += 1
    flush()
    return runs


# ══════════════════════════════════════════════════════════════════════════════
# Documento
# ══════════════════════════════════════════════════════════════════════════════

SKIP_PARA_EXACT = ('\\label', '\\centering', '\\vspace', '\\renewcommand',
                   '\\setlength', '\\pagestyle', '\\thispagestyle',
                   '\\balance', '\\arraystretch', '\\tabcolsep',
                   '\\scriptsize', '\\footnotesize', '\\noindent', '\\small')


def new_par(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10.0,
            before=0, after=3, indent=None, hang=None, line=1.0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if indent is not None:
        pf.left_indent = Inches(indent)
    if hang is not None:
        pf.first_line_indent = Inches(-hang)
    return p


def render_paragraph(doc, text, ctx, size=10.0, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                     after=3):
    stripped = text.strip()
    if not stripped:
        return
    for sk in SKIP_PARA_EXACT:
        if stripped.startswith(sk):
            return
    p = new_par(doc, align=align, size=size, after=after)
    add_runs(p, parse_inline(stripped.replace('\n', ' '), ctx), size=size)
    return p


def add_table_caption(doc, number, caption, ctx):
    """Captions de tabla: 'TABLE I' + texto, arriba, 8pt."""
    p1 = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=8, before=8, after=1)
    p1.paragraph_format.keep_with_next = True
    r = p1.add_run('TABLE %s' % number)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(8)
    if SMALL_CAPS:
        r.font.small_caps = True
    p2 = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=8, after=3)
    p2.paragraph_format.keep_with_next = True
    add_runs(p2, parse_inline(caption, ctx), size=8)


def add_table(doc, label, headers, rows, ctx):
    widths = TABLE_WIDTHS[label]
    fsize = TABLE_FONTS[label]
    total = sum(widths)
    t = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fix_table_layout(t, total)
    set_row_props(t.rows[0], header=True)

    for j, h in enumerate(headers):
        cell = t.cell(0, j)
        set_cell_background(cell, 'EAECEE')
        set_cell_margins(cell)
        pc = cell.paragraphs[0]
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pc.paragraph_format.space_after = Pt(0)
        pc.paragraph_format.line_spacing = 1.0
        add_runs(pc, parse_inline(h, ctx), size=fsize + 0.6)
        for r in pc.runs:
            r.bold = True

    for ri, row in enumerate(rows):
        set_row_props(t.rows[ri + 1], header=False)
        if row and row[0].startswith('\x00SPAN\x00'):
            raw = row[0][len('\x00SPAN\x00'):]
            nc, txt = raw.split('\x00', 1)
            nc = min(int(nc), len(headers))
            merged = t.cell(ri + 1, 0).merge(t.cell(ri + 1, nc - 1))
            set_cell_margins(merged)
            pm = merged.paragraphs[0]
            pm.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pm.paragraph_format.space_before = Pt(2)
            pm.paragraph_format.space_after = Pt(1)
            pm.paragraph_format.line_spacing = 1.0
            add_runs(pm, parse_inline(txt, ctx), size=max(6.5, fsize - 1.2))
            continue
        bg = 'FFFFFF' if ri % 2 == 0 else 'F8F9F9'
        for ci in range(len(headers)):
            cell = t.cell(ri + 1, ci)
            set_cell_background(cell, bg)
            set_cell_margins(cell)
            pc = cell.paragraphs[0]
            pc.alignment = (WD_ALIGN_PARAGRAPH.LEFT if ci == 0
                            else WD_ALIGN_PARAGRAPH.CENTER)
            pc.paragraph_format.space_after = Pt(0)
            pc.paragraph_format.line_spacing = 1.0
            cell_runs = parse_inline(row[ci], ctx) if ci < len(row) else []
            add_runs(pc, cell_runs, size=fsize)
    for row in t.rows:
        for idx, w in enumerate(widths):
            row.cells[idx].width = Inches(w)
    sp = new_par(doc, align=WD_ALIGN_PARAGRAPH.LEFT, size=2, after=2)
    sp.paragraph_format.line_spacing = 1.0
    return t


def add_figure(doc, img_path, caption_runs, ctx, frac, height_cm=None):
    path = os.path.join(BASE, img_path)
    w_target = frac * COLW_NARROW
    w, h = w_target, None
    if height_cm is not None and Image is not None and os.path.exists(path):
        with Image.open(path) as im:
            aspect = im.height / im.width
        h_box = height_cm / 2.54
        if w_target * aspect > h_box:
            h = h_box
            w = h_box / aspect
    p_img = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=10, before=6,
                    after=2)
    p_img.paragraph_format.keep_with_next = True
    if os.path.exists(path):
        run = p_img.add_run()
        if h is None:
            run.add_picture(path, width=Inches(w))
        else:
            run.add_picture(path, width=Inches(w), height=Inches(h))
    else:
        r = p_img.add_run('[figura no encontrada: %s]' % img_path)
        r.font.size = Pt(8)
    cap = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=8.5, after=6)
    add_runs(cap, caption_runs, size=8.5)
    cap.runs[0].bold = True  # "Fig. N."
    for r in cap.runs[1:]:
        r.italic = True


def add_omml_equation(doc, omml_body, eq_num):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = par.paragraph_format
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)
    pf.tab_stops.add_tab_stop(Inches(3.45), WD_TAB_ALIGNMENT.RIGHT)
    xml = ('<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/'
           '2006/math">%s</m:oMath>' % omml_body)
    par._p.append(parse_xml(xml))
    r = par.add_run('\t(%d)' % eq_num)
    r.bold = True
    r.font.name = 'Times New Roman'
    r.font.size = Pt(9.5)


# ══════════════════════════════════════════════════════════════════════════════
# Parseo estructural del .tex
# ══════════════════════════════════════════════════════════════════════════════

def skip_ws_comments(s, i):
    """Avanza sobre espacios y comentarios %... (fin de línea)."""
    n = len(s)
    while i < n:
        c = s[i]
        if c.isspace():
            i += 1
        elif c == '%':
            j = s.find('\n', i)
            i = n if j == -1 else j + 1
        else:
            break
    return i


def strip_comments(chunk):
    """Elimina comentarios %... respetando el % escapado (\\%)."""
    return re.sub(r'(?<!\\)%[^\n]*', '', chunk)


def extract_twocolumn_blocks(src):
    """Extrae los bloques \\twocolumn[{...}] y devuelve (bloques, resto)."""
    blocks = []
    rest = src
    while True:
        m = re.search(r'\\twocolumn\[\{', rest)
        if not m:
            break
        k = m.end() - 1                      # posición de la '{' ya consumida
        content, end = read_group(rest, k)
        end_br = rest.find(']', end)
        blocks.append(content)
        rest = rest[end_br + 1:]
    return blocks, rest


def find_env_end(src, start, env):
    """Índice después de \\end{env} buscando desde `start`."""
    tok = '\\end{' + env + '}'
    j = src.find(tok, start)
    if j == -1:
        raise ValueError('no se encontró %s' % tok)
    return j + len(tok)


def parse_block_header(block, ctx):
    """Parsea título/autores/mentor/afiliación de un bloque \\twocolumn[...]"""
    title = re.search(r'\{\\LARGE\\bfseries (.*?)\\par\}', block, re.S).group(1)
    authors = re.search(r'\{\\normalsize\\bfseries (.*?)\\par\}', block,
                        re.S).group(1)
    m_mentor = re.search(r'\{\\normalsize(\\textit\{.*?)\\par\}', block, re.S)
    mentor = m_mentor.group(1) if m_mentor else None
    affil = re.search(r'\{\\footnotesize (.*?)\\par\}', block, re.S).group(1)
    return title, authors, mentor, affil


def parse_table_env(content, ctx):
    """Extrae caption, label y (headers, rows) de un entorno table."""
    cm = re.search(r'\\caption\{', content)
    caption, _ = read_group(content, cm.end() - 1)
    lm = re.search(r'\\label\{([^}]*)\}', content)
    label = lm.group(1)
    tm = re.search(r'\\begin\{(tabularx|tabular\*)\}', content)
    env = tm.group(1)
    i = tm.end()
    _, i = read_group(content, i)          # ancho objetivo (\textwidth etc.)
    spec, i = read_group(content, i)
    del spec
    eend = content.find('\\end{' + env + '}', i)
    body = content[i:eend]

    rows_raw = []
    for line in body.split('\n'):
        s = line.strip()
        if not s or s in ('\\toprule', '\\midrule', '\\bottomrule'):
            continue
        if s.startswith('\\multicolumn'):
            # fila de nota: \multicolumn{n}{spec}{texto}
            j = len('\\multicolumn')
            cnt, j = read_group(s, j)
            _spec, j = read_group(s, j)
            txt, j = read_group(s, j)
            rows_raw.append(['\x00SPAN\x00' + str(int(cnt)) + '\x00' + txt])
            continue
        if s.endswith('\\\\'):
            s = s[:-2]
        elif '\\\\' in s:
            s = s.split('\\\\')[0]
        cells = [c.strip() for c in s.split('&')]
        rows_raw.append(cells)
    headers = rows_raw[0]
    data = rows_raw[1:]
    return caption, label, headers, data


def parse_item_content(content):
    """Divide contenido de itemize/enumerate en items."""
    parts = re.split(r'\\item\s*', content)
    return [p.strip().replace('\n', ' ') for p in parts if p.strip()]


# ══════════════════════════════════════════════════════════════════════════════
# Construcción principal
# ══════════════════════════════════════════════════════════════════════════════

def build():
    src = open(TEX, encoding='utf-8').read()
    src = src.split('\\begin{document}', 1)[1]
    src = src.split('\\end{document}', 1)[0]

    blocks, body = extract_twocolumn_blocks(src)
    assert len(blocks) == 2, 'se esperaban 2 bloques twocolumn, hay %d' % len(blocks)

    ctx = TexCtx()

    # ── Pre-paseo: numeración de tablas/figuras/ecuaciones y citas ────────────
    tab_i = fig_i = eq_i = 0
    for lm in re.finditer(r'\\label\{([^}]*)\}', src):
        lbl = lm.group(1)
        if lbl.startswith('tab:'):
            tab_i += 1
            ctx.tab_map[lbl] = ROMAN[tab_i - 1]
        elif lbl.startswith('fig:'):
            fig_i += 1
            ctx.fig_map[lbl] = fig_i
        elif lbl.startswith('eq:'):
            eq_i += 1
            ctx.eq_map[lbl] = eq_i
    bib_zone = src.split('\\begin{thebibliography}', 1)[1]
    for k, bm in enumerate(re.finditer(r'\\bibitem\{([^}]*)\}', bib_zone), 1):
        ctx.cite_map[bm.group(1)] = k

    # ── OMML de ecuaciones desde el script anterior ──────────────────────────
    old_src = open(OLD, encoding='utf-8').read()
    omml = {int(m.group(1)): m.group(2).strip()
            for m in re.finditer(r"omml_eq(\d) = '''(.*?)'''", old_src, re.S)}
    assert len(omml) == 5, 'OMML incompleto: %s' % sorted(omml)

    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Times New Roman'
    st.font.size = Pt(10)
    st.font.color.rgb = RGBColor(0, 0, 0)
    st.paragraph_format.line_spacing = 1.0
    st.paragraph_format.space_after = Pt(3)

    # Sección 1 (1 columna): página 1 en inglés
    s1 = doc.sections[0]
    s1.top_margin, s1.bottom_margin = MARGIN_TOP, MARGIN_BOTTOM
    s1.left_margin = s1.right_margin = MARGIN_SIDE
    s1.page_width, s1.page_height = PAGE_W, PAGE_H
    set_columns(s1, 1)

    # Pie de página oficial en todas las páginas
    footer = s1.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    fp.paragraph_format.tab_stops.add_tab_stop(Inches(COLW_FULL),
                                               WD_TAB_ALIGNMENT.RIGHT)
    r = fp.add_run(FOOTER_TEXT + '\t')
    r.font.name = 'Times New Roman'
    r.font.size = Pt(8)
    r.italic = True
    fld = parse_xml(
        '<w:fldSimple %s w:instr=" PAGE ">'
        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" '
        'w:hAnsi="Times New Roman"/><w:sz w:val="16"/></w:rPr>'
        '<w:t>1</w:t></w:r></w:fldSimple>' % nsdecls('w'))
    fp._p.append(fld)

    # ── Página 1: bloque EN ──────────────────────────────────────────────────
    b1 = blocks[0]
    title_en, authors_en, mentor_en, affil_en = parse_block_header(b1, ctx)
    abs_en = re.search(r'\\begin\{abstract\}(.*?)\\end\{abstract\}', b1,
                       re.S).group(1)
    kw_en = re.search(r'\\begin\{IEEEkeywords\}(.*?)\\end\{IEEEkeywords\}',
                      b1, re.S).group(1)

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=8, after=8)
    rr = p.add_run(DOI_TEXT)
    rr.font.size = Pt(8)
    rr.font.name = 'Times New Roman'

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=18, before=6,
                after=10)
    add_runs(p, parse_inline(title_en, ctx), size=18)
    for r in p.runs:
        r.bold = True

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=11, after=4 if mentor_en else 6)
    add_runs(p, parse_inline(authors_en, ctx), size=11)
    for r in p.runs:
        r.bold = True

    if mentor_en:
        p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=10, after=6)
        add_runs(p, parse_inline(mentor_en, ctx), size=10)

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=9, after=14)
    add_runs(p, parse_inline(affil_en, ctx), size=9)

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=9, after=6)
    r = p.add_run('Abstract\u2014')
    r.bold = True
    r.italic = True
    r.font.size = Pt(9)
    add_runs(p, parse_inline(abs_en.strip().replace('\n', ' '), ctx), size=9)

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=9, after=6)
    r = p.add_run('Keywords\u2014')
    r.bold = True
    r.italic = True
    r.font.size = Pt(9)
    add_runs(p, parse_inline(kw_en.strip().replace('\n', ' '), ctx), size=9)

    # ── Página 2: título ES (1 columna) ──────────────────────────────────────
    add_section(doc, WD_SECTION.NEW_PAGE, 1)
    b2 = blocks[1]
    title_es, authors_es, mentor_es, affil_es = parse_block_header(b2, ctx)

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=18, before=6,
                after=10)
    add_runs(p, parse_inline(title_es, ctx), size=18)
    for r in p.runs:
        r.bold = True

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=11, after=4 if mentor_es else 6)
    add_runs(p, parse_inline(authors_es, ctx), size=11)
    for r in p.runs:
        r.bold = True

    if mentor_es:
        p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=10, after=6)
        add_runs(p, parse_inline(mentor_es, ctx), size=10)

    p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=9, after=14)
    add_runs(p, parse_inline(affil_es, ctx), size=9)

    # ── Cuerpo a 2 columnas ──────────────────────────────────────────────────
    add_section(doc, WD_SECTION.CONTINUOUS, 2)

    state = {'sec': 0, 'sub': 0}

    def heading_section(name, star=False):
        state['sub'] = 0
        if star:
            txt = name
        else:
            state['sec'] += 1
            txt = '%s. %s' % (ROMAN[state['sec'] - 1], name)
        p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=10, before=10,
                    after=4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(txt)
        r.font.size = Pt(10)
        if SMALL_CAPS:
            r.font.small_caps = True
        else:
            r.text = txt.upper()

    def heading_sub(name):
        state['sub'] += 1
        p = new_par(doc, align=WD_ALIGN_PARAGRAPH.LEFT, size=10, before=6,
                    after=3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run('%s. %s' % (LETTERS[state['sub'] - 1], name))
        r.font.size = Pt(10)
        r.italic = True

    def render_text(chunk):
        for para in re.split(r'\n\s*\n', strip_comments(chunk)):
            render_paragraph(doc, para, ctx)

    MARK = re.compile(
        r'\\section\*?\{|\\subsection\{|'
        r'\\begin\{(table\*?|figure\*?|equation|itemize|enumerate|'
        r'abstract|IEEEkeywords|thebibliography)\}|\\balance\b')

    pos = 0
    while pos < len(body):
        m = MARK.search(body, pos)
        if not m:
            render_text(body[pos:])
            break
        render_text(body[pos:m.start()])
        tok = m.group(0)
        if tok.startswith('\\balance'):
            pos = m.end()
            continue
        if tok.startswith('\\section') or tok.startswith('\\subsection'):
            name, pos = read_group(body, m.end() - 1)
            star = tok.startswith('\\section*')
            if tok.startswith('\\subsection'):
                heading_sub(name)
            else:
                heading_section(name, star)
            continue
        env = m.group(1)
        begin_len = len('\\begin{' + env + '}')
        end = find_env_end(body, m.end() - begin_len, env)
        content = body[m.end():end - len('\\end{' + env + '}')]
        pos = end

        if env in ('table', 'table*'):
            caption, label, headers, rows = parse_table_env(content, ctx)
            if env.endswith('*'):
                add_section(doc, WD_SECTION.CONTINUOUS, 1)
            add_table_caption(doc, ctx.tab_map[label], caption, ctx)
            add_table(doc, label, headers, rows, ctx)
            if env.endswith('*'):
                add_section(doc, WD_SECTION.CONTINUOUS, 2)
        elif env in ('figure', 'figure*'):
            im = re.search(r'\\includegraphics\[([^\]]*)\]\{([^}]*)\}', content)
            opts, img = im.group(1), im.group(2)
            if '\\textwidth' in opts:
                frac = COLW_FULL / COLW_NARROW
            else:
                wm = re.search(r'width=([\d.]+)?\\columnwidth', opts)
                frac = float(wm.group(1)) if (wm and wm.group(1)) else 1.0
            hm = re.search(r'height=([\d.]+)\s*cm', opts)
            keep = 'keepaspectratio' in opts
            capm = re.search(r'\\caption\{', content)
            caption, _ = read_group(content, capm.end() - 1)
            lblm = re.search(r'\\label\{([^}]*)\}', content)
            fig_no = ctx.fig_map[lblm.group(1)]
            if env == 'figure*':
                add_section(doc, WD_SECTION.CONTINUOUS, 1)
            add_figure(doc, img,
                       [{'t': 'Fig. %d. ' % fig_no}] +
                       parse_inline(caption, ctx),
                       ctx, frac,
                       height_cm=float(hm.group(1)) if (hm and keep) else None)
            if env == 'figure*':
                add_section(doc, WD_SECTION.CONTINUOUS, 2)
        elif env == 'equation':
            lblm = re.search(r'\\label\{([^}]*)\}', content)
            eq_no = ctx.eq_map[lblm.group(1)]
            add_omml_equation(doc, omml[eq_no], eq_no)
        elif env == 'itemize':
            for item in parse_item_content(content):
                p = new_par(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10,
                            after=3, indent=0.18, hang=0.12)
                add_runs(p, [{'t': '\u2022\u00a0'}] + parse_inline(item, ctx),
                         size=10)
        elif env == 'enumerate':
            for k, item in enumerate(parse_item_content(content), 1):
                p = new_par(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=10,
                            after=3, indent=0.22, hang=0.16)
                add_runs(p, [{'t': '%d)\u00a0' % k}] + parse_inline(item, ctx),
                         size=10)
        elif env in ('abstract', 'IEEEkeywords'):
            lead = 'Resumen' if env == 'abstract' else 'Palabras clave'
            p = new_par(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=9, after=6)
            r = p.add_run('%s\u2014' % lead)
            r.bold = True
            r.italic = True
            r.font.size = Pt(9)
            add_runs(p, parse_inline(content.strip().replace('\n', ' '), ctx),
                     size=9)
        elif env == 'thebibliography':
            p = new_par(doc, align=WD_ALIGN_PARAGRAPH.CENTER, size=10,
                        before=10, after=4)
            r = p.add_run('Referencias')
            r.font.size = Pt(10)
            if SMALL_CAPS:
                r.font.small_caps = True
            else:
                r.text = 'REFERENCIAS'
            keys = re.findall(r'\\bibitem\{([^}]*)\}', content)
            items = re.split(r'\\bibitem\{[^}]*\}', content)[1:]
            for k, (key, item) in enumerate(zip(keys, items), 1):
                del key
                p = new_par(doc, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=9,
                            after=2, indent=0.25, hang=0.25)
                add_runs(p, [{'t': '[%d]\u00a0' % k}] +
                         parse_inline(item.strip().replace('\n', ' '), ctx),
                         size=9)

    doc.save(OUT)
    print('OK -> %s' % OUT)


if __name__ == '__main__':
    build()
