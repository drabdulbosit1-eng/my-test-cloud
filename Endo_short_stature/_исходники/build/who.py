"""Центильные кривые по нормам ВОЗ: рост для возраста и ИМТ для возраста.

Данные: стандарты ВОЗ 0–5 лет (2006) и референс ВОЗ 5–19 лет (2007), параметры L, M, S по месяцам.
Файл данных: _исходники/данные/ВОЗ/who_lms.json
  {"hfa": {"M": [[месяц, L, M, S], ...], "F": [...]}, "bfa": {...}}
Значение для СО z: X = M * (1 + L*S*z)^(1/L), при L = 0: X = M * exp(S*z).

Точки ребёнка задаются в СО и пересчитываются по тем же таблицам: это пример на настоящих нормах.
"""
import json
import math
import os

from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from growth import Panel, _label, _line, _poly, _rgb
from style import C

DATA = os.path.join(os.path.dirname(__file__), '..', 'данные', 'ВОЗ', 'who_lms.json')
_LMS = None

# центили на бланках ВОЗ и их СО
CENTILES = [(97, 1.881), (85, 1.036), (50, 0.0), (15, -1.036), (3, -1.881)]


def available():
    return os.path.exists(DATA)


def _lms():
    global _LMS
    if _LMS is None:
        _LMS = json.load(open(DATA, encoding='utf-8'))
    return _LMS


def lms(kind, sex, age_years):
    rows = _lms()[kind][sex]
    m = age_years * 12.0
    lo = None
    for r in rows:
        if r[0] <= m:
            lo = r
        else:
            hi = r
            if lo is None:
                return r[1:]
            f = (m - lo[0]) / (hi[0] - lo[0])
            return [lo[i] + (hi[i] - lo[i]) * f for i in (1, 2, 3)]
    return rows[-1][1:]


def value(kind, sex, t, z):
    L, M, S = lms(kind, sex, t)
    if abs(L) < 1e-6:
        return M * math.exp(S * z)
    return M * (1 + L * S * z) ** (1 / L)


def _steps(t0, t1, n=80):
    return [t0 + (t1 - t0) * i / n for i in range(n + 1)]


def panel(shapes, x, y, w, h, sex, t0, t1, kind, patient, events=(), title=None, grid_step=None,
          label_step=None, unit=''):
    """Один график как на бланке ВОЗ: сетка, пять центилей, точки примера."""
    ts = _steps(t0, t1)
    lo = [value(kind, sex, t, -2.6) for t in ts]
    hi = [value(kind, sex, t, 2.4) for t in ts]
    pv = [value(kind, sex, a, z) for a, z in patient]
    vmin, vmax = min(lo + pv), max(hi + pv)
    g = grid_step or (5 if kind == 'hfa' else 1)
    ls = label_step or (10 if kind == 'hfa' else 2)
    v0 = math.floor(vmin / g) * g
    v1 = math.ceil(vmax / g) * g
    lx = Inches(0.42)                  # место под подписи оси слева
    rx = Inches(0.32)                  # место под номера центилей справа
    P = Panel(x + lx, y, w - lx - rx, h, t0, t1, v0, v1)

    # сетка
    v = v0
    while v <= v1 + 1e-6:
        major = abs((v / ls) - round(v / ls)) < 1e-6
        _line(shapes, [(P.x, P.py(v)), (P.x + P.w, P.py(v))], C['rule'] if major else 'EEF2F4',
              0.75 if major else 0.5, name=f'{kind}: сетка {v:g}')
        if major:
            _label(shapes, x, P.py(v) - Emu(80000), lx - Emu(50000), Emu(160000), f'{v:g}', size=9,
                   align=PP_ALIGN.RIGHT)
        v += g
    a = math.ceil(t0)
    while a <= t1 + 1e-6:
        _line(shapes, [(P.px(a), P.y), (P.px(a), P.y + P.h)], 'EEF2F4', 0.5, name=f'{kind}: год {a}')
        step = 1 if (t1 - t0) <= 10 else 2
        if (a - math.ceil(t0)) % step == 0:
            _label(shapes, P.px(a) - Emu(150000), P.y + P.h + Emu(30000), Emu(300000), Emu(170000), str(int(a)),
                   size=9, align=PP_ALIGN.CENTER)
        a += 1
    _line(shapes, [(P.x, P.y + P.h), (P.x + P.w, P.y + P.h)], C['axis'], 0.75, name=f'{kind}: ось')

    # полоса 3–97 и центили
    z3, z97 = -1.881, 1.881
    band = [(P.px(t), P.py(value(kind, sex, t, z97))) for t in ts] + \
           [(P.px(t), P.py(value(kind, sex, t, z3))) for t in reversed(ts)]
    _poly(shapes, band, C['band'], name=f'{kind}: полоса 3–97')
    for c, z in CENTILES:
        pts = [(P.px(t), P.py(value(kind, sex, t, z))) for t in ts]
        col = C['axis'] if c == 50 else C['centile']
        _line(shapes, pts, col, 1.25 if c == 50 else 0.75, name=f'{kind}: {c}-й центиль')
        _label(shapes, P.x + P.w + Emu(40000), P.py(value(kind, sex, t1, z)) - Emu(75000), rx, Emu(150000),
               str(c), size=9)

    # пример ребёнка
    pts = []
    for (a0, z0), (a1, z1) in zip(patient, patient[1:]):
        for t in _steps(a0, a1, 12)[:-1]:
            zz = z0 + (z1 - z0) * (t - a0) / (a1 - a0)
            pts.append((P.px(t), P.py(value(kind, sex, t, zz))))
    a, z = patient[-1]
    pts.append((P.px(a), P.py(value(kind, sex, a, z))))
    col = C['height'] if kind == 'hfa' else C['weight']
    _line(shapes, pts, col, 2.5, name=f'{kind}: пример')
    for a, z in patient:
        d = Emu(58000)
        dot = shapes.add_shape(9, P.px(a) - d, P.py(value(kind, sex, a, z)) - d, 2 * d, 2 * d)
        dot.fill.solid()
        dot.fill.fore_color.rgb = _rgb(col)
        dot.line.color.rgb = _rgb('FFFFFF')
        dot.line.width = Pt(0.75)
        dot.shadow.inherit = False
        from diagrams import _no_effect
        _no_effect(dot)
        dot.name = f'{kind}: точка {a}'

    # события
    for i, ev in enumerate(events):
        a, text = ev[0], ev[1]
        _line(shapes, [(P.px(a), P.y), (P.px(a), P.y + P.h)], C['muted'], 0.75, dash=MSO_LINE_DASH_STYLE.DASH,
              name=f'событие {a}')
        lw = Emu(2350000)
        bg_y = P.y + Emu(20000) + i * Emu(195000)
        if P.px(a) + Emu(50000) + lw <= P.x + P.w:
            tb = _label(shapes, P.px(a) + Emu(50000), bg_y, lw, Emu(185000), text, size=10, color=C['text'],
                        anchor=MSO_ANCHOR.TOP)
        else:
            tb = _label(shapes, P.px(a) - Emu(50000) - lw, bg_y, lw, Emu(185000), text, size=10, color=C['text'],
                        anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.RIGHT)
    if title:
        _label(shapes, x, y - Emu(250000), w, Emu(220000), title, size=12, color=C['text'], bold=True,
               anchor=MSO_ANCHOR.BOTTOM)
    if unit:
        _label(shapes, x, y - Emu(10000), lx - Emu(60000), Emu(150000), unit, size=9, align=PP_ALIGN.RIGHT)
    return P


SEX_WORD = {'M': 'мальчики', 'F': 'девочки'}


def chart(shapes, x, y, w, h, sex, t0, t1, hz, bz, events=(), title=None):
    """Рост для возраста (крупно) и ИМТ для возраста (ниже), обе по ВОЗ."""
    gap = Inches(0.55)
    hh = (h - gap) * 0.66
    bh = h - gap - hh
    head = (title + ': ' if title else '') + f'рост для возраста, {SEX_WORD[sex]}, см'
    panel(shapes, x, y, w, hh, sex, t0, t1, 'hfa', hz, events, title=head)
    panel(shapes, x, y + hh + gap, w, bh, sex, t0, t1, 'bfa', bz, (), title=f'ИМТ для возраста, кг/м²',
          grid_step=1, label_step=2)
