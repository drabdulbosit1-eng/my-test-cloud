"""Сборка лекции из JSON-файлов текста.

python3 build_deck.py <папка_с_json> <выход.pptx>
"""
import copy
import glob
import json
import os
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

sys.path.insert(0, os.path.dirname(__file__))
from style import (BODY_BOTTOM, BODY_Y, C, CONTENT_W, EVID_Y, FONT_BODY, FONT_HEAD, MX, PARTS, RULER_Y,
                   SIZE_BODY, SIZE_EVID, SIZE_SRC, SIZE_TABLE, SIZE_TITLE, SLIDE_H, SLIDE_W, SRC_Y,
                   TITLE_H, TITLE_Y, TOTAL)
import growth
from figures import FIGS

NB = ' '
UNITS = (r'мг|мкг|мл|кг|см|мм|м²|лет|года|год|мес|нед|дн|сут|ч|нмоль/л|пмоль/л|нг/мл|пг/мл|мЕд/л|МЕ/л|'
         r'мкг/дл|мг/л|%|Гр|кДа|раз|случа|детей|ребён|исследован|пациент|мальчик|девоч|человек|из|СО')


def typo(s):
    """Неразрывные пробелы: число и единица, знаки сравнения, n = 24."""
    if not s:
        return s
    s = re.sub(r'(\d) (?=(' + UNITS + r'))', r'\1' + NB, s)
    s = re.sub(r'([<>≤≥=±−]) (?=[\d−])', r'\1' + NB, s)
    s = re.sub(r'(\bn|\bр|\bp) ([=<>])', r'\1' + NB + r'\2', s)
    s = re.sub(r'(\d) (–|−) (\d)', r'\1\2\3', s)
    s = s.replace('св. Т4', 'св.' + NB + 'Т4')
    return s


def rgb(h):
    return RGBColor.from_string(h)


def add_runs(par, text, size, color, font=FONT_BODY, bold=False):
    """Текст с курсивом для *генов*."""
    parts = re.split(r'(\*[^*]+\*)', typo(text))
    for part in parts:
        if not part:
            continue
        r = par.add_run()
        it = part.startswith('*') and part.endswith('*')
        r.text = part[1:-1] if it else part
        f = r.font
        f.size = size
        f.name = font
        f.bold = bold
        f.italic = it
        f.color.rgb = rgb(color)


def textbox(shapes, x, y, w, h, anchor=MSO_ANCHOR.TOP, name=None):
    tb = shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    if name:
        tb.name = name
    return tb, tf


def set_bullet(par, level_indent=Inches(0.28), char='–'):
    pPr = par._p.get_or_add_pPr()
    pPr.set('marL', str(int(level_indent)))
    pPr.set('indent', str(-int(level_indent)))
    for tag in ('a:buNone', 'a:buChar', 'a:buAutoNum'):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn('a:buClr'))
    srgb = etree.SubElement(buClr, qn('a:srgbClr'))
    srgb.set('val', C['height'])
    bu = etree.SubElement(pPr, qn('a:buChar'))
    bu.set('char', char)


def space_after(par, pt):
    par.space_after = Pt(pt)


def title(slide, text, dark=False):
    tb, tf = textbox(slide.shapes, MX, TITLE_Y, CONTENT_W, TITLE_H, anchor=MSO_ANCHOR.TOP, name='Заголовок')
    p = tf.paragraphs[0]
    p.line_spacing = 0.95
    add_runs(p, text, SIZE_TITLE, C['onink'] if dark else C['ink'], font=FONT_HEAD, bold=True)


def bullets(slide, items, x, y, w, h, size=SIZE_BODY, color=None, gap=10):
    tb, tf = textbox(slide.shapes, x, y, w, h, name='Текст')
    first = True
    for it in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        set_bullet(p)
        space_after(p, gap)
        p.line_spacing = 1.0
        add_runs(p, it, size, color or C['text'])
    return tb


def footer(slide, n, evidence, source, dark=False):
    col = C['onink2'] if dark else C['muted']
    if evidence:
        tb, tf = textbox(slide.shapes, MX, EVID_Y, CONTENT_W, Inches(0.26), name='Уровень доказательности')
        add_runs(tf.paragraphs[0], evidence, SIZE_EVID, col)
        tf.paragraphs[0].runs[0].font.italic = True if tf.paragraphs[0].runs else None
    if source:
        tb, tf = textbox(slide.shapes, MX, SRC_Y, CONTENT_W, Inches(0.42), name='Источник')
        add_runs(tf.paragraphs[0], source, SIZE_SRC, col)
    ruler(slide, n, dark)


def ruler(slide, n, dark=False):
    """Шкала-ростомер: где слушатель находится в лекции."""
    sh = slide.shapes
    x0, x1 = int(MX), int(SLIDE_W - MX)
    y = int(RULER_Y + Inches(0.05))
    base_col = C['onink2'] if dark else C['rule']
    fill_col = C['onink'] if dark else C['height']

    def rect(x, yy, w, h, col, name):
        r = sh.add_shape(1, int(x), int(yy), max(int(w), 1), int(h))
        r.fill.solid()
        r.fill.fore_color.rgb = rgb(col)
        r.line.fill.background()
        r.shadow.inherit = False
        r.name = name
        return r

    step = (x1 - x0) / TOTAL
    rect(x0, y, x1 - x0, Inches(0.03), base_col, 'Ростомер: шкала')
    rect(x0, y - Inches(0.015), step * n, Inches(0.06), fill_col, 'Ростомер: пройдено')
    for i in range(TOTAL + 1):
        major = any(i == a - 1 for a, b, _ in PARTS) or i == TOTAL
        hgt = Inches(0.16) if major else Inches(0.07)
        rect(x0 + step * i - Emu(4000), y - hgt, Emu(8000), hgt, base_col, 'Ростомер: деление')
    for a, b, name in PARTS:
        cx = x0 + step * (a - 1)
        w = step * (b - a + 1)
        cur = a <= n <= b
        tb, tf = textbox(sh, cx, y + Inches(0.07), w, Inches(0.22), name='Ростомер: раздел')
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        col = (C['onink'] if dark else C['ink']) if cur else (C['onink2'] if dark else C['muted'])
        add_runs(p, name, Pt(10), col, bold=cur)
    tb, tf = textbox(sh, x1 - Inches(0.6), y - Inches(0.36), Inches(0.6), Inches(0.24), name='Номер слайда')
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    add_runs(p, f'{n}', Pt(11), C['onink2'] if dark else C['muted'])


def background(slide, col):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = rgb(col)


def table(slide, spec, x, y, w, h, size=SIZE_TABLE, col_w=None):
    header, rows = spec['header'], spec['rows']
    nr, nc = len(rows) + 1, len(header)
    gt = slide.shapes.add_table(nr, nc, int(x), int(y), int(w), int(h))
    gt.name = 'Таблица'
    tbl = gt.table
    # без встроенного стиля
    tblPr = tbl._tbl.tblPr
    tblPr.set('firstRow', '1')
    tblPr.set('bandRow', '0')
    style = tblPr.find(qn('a:tableStyleId'))
    if style is None:
        style = etree.SubElement(tblPr, qn('a:tableStyleId'))
    style.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'  # No Style, Table Grid
    if col_w:
        tot = sum(col_w)
        for i, cw in enumerate(col_w):
            tbl.columns[i].width = int(w * cw / tot)
    for r in range(nr):
        for c in range(nc):
            cell = tbl.cell(r, c)
            txt = header[c] if r == 0 else (rows[r - 1][c] if c < len(rows[r - 1]) else '')
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = Inches(0.04)
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            for run in list(p.runs):
                run._r.getparent().remove(run._r)
            add_runs(p, str(txt), size, C['onink'] if r == 0 else C['text'], bold=(r == 0))
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(C['ink'] if r == 0 else (C['paper2'] if r % 2 == 0 else C['paper']))
            # тонкие линии
            tcPr = cell._tc.get_or_add_tcPr()
            for side in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
                ln = tcPr.find(qn(side))
                if ln is None:
                    ln = etree.SubElement(tcPr, qn(side))
                ln.set('w', '6350')
                for ch in list(ln):
                    ln.remove(ch)
                sf = etree.SubElement(ln, qn('a:solidFill'))
                cl = etree.SubElement(sf, qn('a:srgbClr'))
                cl.set('val', C['rule'])
    return gt


def notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


def caption(slide, x, y, w, text='схема, а не данные конкретного пациента'):
    tb, tf = textbox(slide.shapes, x, y, w, Inches(0.24), name='Подпись к схеме')
    add_runs(tf.paragraphs[0], text, Pt(11), C['muted'])
    tf.paragraphs[0].runs[0].font.italic = True


def legend(slide, x, y):
    for i, (col, txt) in enumerate(((C['height'], 'пациент: рост'), (C['weight'], 'пациент: масса'),
                                    (C['band'], 'норма: 3–97-й центиль'))):
        r = slide.shapes.add_shape(1, int(x + Inches(2.2) * i), int(y + Inches(0.07)), int(Inches(0.3)),
                                   int(Inches(0.1)))
        r.fill.solid()
        r.fill.fore_color.rgb = rgb(col)
        r.line.fill.background()
        r.shadow.inherit = False
        r.name = 'Легенда'
        tb, tf = textbox(slide.shapes, x + Inches(2.2) * i + Inches(0.38), y, Inches(1.8), Inches(0.24),
                         name='Легенда')
        add_runs(tf.paragraphs[0], txt, Pt(11), C['muted'])


# ---------- типы слайдов ----------

def s_title(prs, d):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, C['ink'])
    sh = s.shapes
    # мотив: одна центильная кривая
    P = growth.Panel(int(Inches(0.0)), int(Inches(3.9)), int(SLIDE_W), int(Inches(2.6)), 2, 18, 80, 185)
    for z, col, wpt in ((1.88, C['onink2'], 0.75), (0, C['onink2'], 0.75), (-1.88, C['onink2'], 0.75)):
        pts = [(P.px(t), P.py(growth.height('M', t, z))) for t in [2 + i * 16 / 60 for i in range(61)]]
        ln = growth._line(sh, pts, col, wpt, name='Мотив: центиль')
    pts = []
    for t in [2 + i * 16 / 60 for i in range(61)]:
        z = -0.2 if t < 9 else (-0.2 - (t - 9) * 0.35 if t < 13 else -1.6 + (t - 13) * 0.3)
        pts.append((P.px(t), P.py(growth.height('M', t, z))))
    growth._line(sh, pts, C['weight'], 2.5, name='Мотив: пациент')
    tb, tf = textbox(sh, MX, Inches(1.0), CONTENT_W, Inches(1.9), anchor=MSO_ANCHOR.BOTTOM, name='Название')
    add_runs(tf.paragraphs[0], d['title'], Pt(40), C['onink'], font=FONT_HEAD, bold=True)
    body = d.get('body') or []
    tb, tf = textbox(sh, MX, Inches(3.05), CONTENT_W, Inches(1.0), name='Подзаголовок')
    for i, line in enumerate(body):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        add_runs(p, line, Pt(18), C['onink2'])
    notes(s, d.get('notes'))
    return s


def s_dark(prs, d):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, C['ink'])
    title(s, d['title'], dark=True)
    items = d.get('body') or []
    bullets(s, items, MX, BODY_Y, CONTENT_W, BODY_BOTTOM - BODY_Y, size=Pt(20), color=C['onink'], gap=12)
    footer(s, d['n'], d.get('evidence'), d.get('source'), dark=True)
    notes(s, d.get('notes'))
    return s


def s_content(prs, d):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, C['paper'])
    title(s, d['title'])
    n = d['n']
    fig = FIGS.get(n)
    items = d.get('body') or []
    tbl = d.get('table')
    if tbl and not tbl.get('rows'):
        tbl = None
    top, bottom = BODY_Y, BODY_BOTTOM
    if fig and fig['type'] == 'growth':
        fx, fw = MX, Inches(7.4)
        growth.growth_chart(s.shapes, fx, top, fw, bottom - top - Inches(0.85), fig['sex'], fig['t0'], fig['t1'],
                            fig['hz'], fig['wz'], fig.get('events', ()))
        legend(s, fx, bottom - Inches(0.52))
        caption(s, fx, bottom - Inches(0.25), fw)
        bullets(s, items, MX + Inches(7.75), top + Inches(0.2), CONTENT_W - Inches(7.75), bottom - top - Inches(0.2),
                size=Pt(18))
    elif fig and fig['type'] == 'two_growth':
        half = (CONTENT_W - Inches(0.5)) / 2
        h = bottom - top - Inches(0.62) - (Inches(0.55) if items else 0)
        for i, g in enumerate(fig['charts']):
            growth.growth_chart(s.shapes, MX + i * (half + Inches(0.5)), top, half, h, g['sex'], g['t0'], g['t1'],
                                g['hz'], g['wz'], g.get('events', ()), title=g['title'])
        legend(s, MX, top + h + Inches(0.27))
        caption(s, MX + Inches(6.9), top + h + Inches(0.27), Inches(5.2))
        if items:
            bullets(s, items, MX, top + h + Inches(0.62), CONTENT_W, Inches(0.6), size=Pt(17), gap=4)
    elif fig and fig['type'] == 'thumbs':
        cols, rows_ = fig.get('cols', 3), fig.get('rows', 2)
        gw, gh = Inches(0.35), Inches(0.25)
        tw = (CONTENT_W - gw * (cols - 1)) / cols
        th = (bottom - top - Inches(0.35) - gh * (rows_ - 1)) / rows_
        for i, t in enumerate(fig['items']):
            r, c = divmod(i, cols)
            growth.z_thumb(s.shapes, MX + c * (tw + gw), top + r * (th + gh), tw, th, t['title'], t['t0'], t['t1'],
                           t['hz'], t['wz'], t.get('note'))
        # легенда
        for i, (col, txt) in enumerate(((C['height'], 'рост, СО'), (C['weight'], 'масса, СО'),
                                        (C['band'], 'от −2 до +2 СО'))):
            rr = s.shapes.add_shape(1, int(MX + Inches(2.0) * i), int(bottom - Inches(0.22)), int(Inches(0.3)),
                                    int(Inches(0.1)))
            rr.fill.solid()
            rr.fill.fore_color.rgb = rgb(col)
            rr.line.fill.background()
            rr.shadow.inherit = False
            tb, tf = textbox(s.shapes, MX + Inches(2.0) * i + Inches(0.38), bottom - Inches(0.29), Inches(1.6),
                             Inches(0.24), name='Легенда')
            add_runs(tf.paragraphs[0], txt, Pt(11), C['muted'])
        caption(s, MX + Inches(6.4), bottom - Inches(0.29), Inches(5.7))
    elif tbl and items:
        nrows = len(tbl['rows']) + 1
        th = min(bottom - top - Inches(1.2), Inches(0.5) * nrows + Inches(0.3))
        table(s, tbl, MX, top, CONTENT_W, th, col_w=d.get('col_w'))
        bullets(s, items, MX, top + th + Inches(0.3), CONTENT_W, bottom - top - th - Inches(0.3), size=Pt(17), gap=4)
    elif tbl:
        nrows = len(tbl['rows']) + 1
        th = min(bottom - top, Inches(0.62) * nrows)
        table(s, tbl, MX, top, CONTENT_W, th, col_w=d.get('col_w'))
    else:
        call = d.get('callout')
        if call:
            cw = Inches(3.6)
            tb, tf = textbox(s.shapes, MX, top + Inches(0.2), cw, Inches(1.4), anchor=MSO_ANCHOR.BOTTOM, name='Число')
            add_runs(tf.paragraphs[0], call['value'], Pt(54), C['height'], font=FONT_HEAD, bold=True)
            tb, tf = textbox(s.shapes, MX, top + Inches(1.7), cw, Inches(1.8), name='Подпись к числу')
            add_runs(tf.paragraphs[0], call['label'], Pt(16), C['muted'])
            bullets(s, items, MX + cw + Inches(0.5), top + Inches(0.2), CONTENT_W - cw - Inches(0.5),
                    bottom - top - Inches(0.2))
        else:
            bullets(s, items, MX, top + Inches(0.15), Inches(10.5), bottom - top - Inches(0.15))
    footer(s, n, d.get('evidence'), d.get('source'))
    notes(s, d.get('notes'))
    return s


def s_refs(prs, d):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    background(s, C['paper'])
    title(s, d['title'])
    items = d.get('body') or []
    half = (len(items) + 1) // 2
    colw = (CONTENT_W - Inches(0.4)) / 2
    for i, chunk in enumerate((items[:half], items[half:])):
        tb, tf = textbox(s.shapes, MX + i * (colw + Inches(0.4)), BODY_Y - Inches(0.2), colw,
                         RULER_Y - BODY_Y - Inches(0.05), name='Литература')
        for j, ref in enumerate(chunk):
            p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
            p.space_after = Pt(1)
            add_runs(p, ref, Pt(8), C['text'])
    ruler(s, d['n'])
    notes(s, d.get('notes'))
    return s


def load(folder):
    slides = []
    for f in sorted(glob.glob(os.path.join(folder, '*.json'))):
        with open(f, encoding='utf-8') as fh:
            slides.extend(json.load(fh))
    slides.sort(key=lambda d: d['n'])
    return slides


def build(folder, out):
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    for d in load(folder):
        lay = d.get('layout')
        if d['n'] == 1 or lay == 'title':
            s_title(prs, d)
        elif lay == 'references':
            s_refs(prs, d)
        elif lay == 'conclusions':
            s_dark(prs, d)
        else:
            s_content(prs, d)
    prs.core_properties.title = 'Эндокринные причины низкорослости у детей и подростков'
    prs.core_properties.language = 'ru-RU'
    prs.save(out)
    print('saved', out)


if __name__ == '__main__':
    build(sys.argv[1], sys.argv[2])
