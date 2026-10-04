"""Helpers for restyling the finerenone deck with python-pptx + lxml."""
import copy
import re

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

TITLE_FONT = "Arial Narrow"
BODY_FONT = "Arial Narrow"
PURPLE = "7030A0"
BLACK = "000000"
GREY = "595959"
GREEN, GREEN_T = "2E7D4F", "EAF4EE"
AMBER, AMBER_T = "A76A10", "FBF1E1"
RED, RED_T = "B5323C", "FBEAEA"
TEAL, TEAL_T = "007F8D", "E6F3F4"
PURPLE_T = "F3EDF8"
GREY_T = "F2F2F2"
BLUE = "156082"

# order of child elements inside a:rPr (CT_TextCharacterProperties)
RPR_ORDER = ["ln", "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
             "effectLst", "effectDag", "highlight", "uLnTx", "uLn", "uFillTx", "uFill",
             "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"]
FILLS = {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}


def _local(el):
    return etree.QName(el).localname


def _sort_rpr(rpr):
    kids = list(rpr)
    kids.sort(key=lambda e: RPR_ORDER.index(_local(e)) if _local(e) in RPR_ORDER else 99)
    for k in kids:
        rpr.remove(k)
    for k in kids:
        rpr.append(k)


def style_rpr(rpr, size=None, font=None, color=None, bold=None, italic=None,
              drop_links=True, drop_underline=False, drop_highlight=False):
    """Set properties on an a:rPr / a:endParaRPr / a:defRPr element in place."""
    if size is not None:
        rpr.set("sz", str(int(round(size * 100))))
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    if italic is not None:
        rpr.set("i", "1" if italic else "0")
    if drop_underline and rpr.get("u"):
        del rpr.attrib["u"]
    if color is not None:
        for k in list(rpr):
            if _local(k) in FILLS:
                rpr.remove(k)
        sf = etree.SubElement(rpr, qn("a:solidFill"))
        c = etree.SubElement(sf, qn("a:srgbClr"))
        c.set("val", color)
    if font is not None:
        for k in list(rpr):
            if _local(k) in ("latin", "cs"):
                rpr.remove(k)
        lat = etree.SubElement(rpr, qn("a:latin"))
        lat.set("typeface", font)
        cs = etree.SubElement(rpr, qn("a:cs"))
        cs.set("typeface", font)
    if drop_links:
        for k in list(rpr):
            if _local(k) in ("hlinkClick", "hlinkMouseOver"):
                rpr.remove(k)
    if drop_highlight:
        for k in list(rpr):
            if _local(k) == "highlight":
                rpr.remove(k)
    _sort_rpr(rpr)


def get_rpr(r):
    rpr = r.find(qn("a:rPr"))
    if rpr is None:
        rpr = etree.Element(qn("a:rPr"))
        r.insert(0, rpr)
    return rpr


def effective_size(r):
    """Size of a run in pt: its own sz, else the paragraph defRPr sz, else None."""
    rpr = r.find(qn("a:rPr"))
    if rpr is not None and rpr.get("sz"):
        return int(rpr.get("sz")) / 100
    p = r.getparent()
    ppr = p.find(qn("a:pPr"))
    if ppr is not None:
        d = ppr.find(qn("a:defRPr"))
        if d is not None and d.get("sz"):
            return int(d.get("sz")) / 100
    return None


def run_color(r):
    rpr = r.find(qn("a:rPr"))
    if rpr is None:
        p = r.getparent().find(qn("a:pPr"))
        rpr = p.find(qn("a:defRPr")) if p is not None else None
    if rpr is None:
        return None
    c = rpr.find("a:solidFill/a:srgbClr", NS)
    return c.get("val") if c is not None else None


def materialize_defrpr(p_el):
    """Copy paragraph-level a:pPr/a:defRPr into every run that lacks explicit values."""
    ppr = p_el.find(qn("a:pPr"))
    if ppr is None:
        return
    d = ppr.find(qn("a:defRPr"))
    if d is None:
        return
    for r in p_el.findall(qn("a:r")):
        rpr = get_rpr(r)
        for attr in ("sz", "b", "i"):
            if d.get(attr) is not None and rpr.get(attr) is None:
                rpr.set(attr, d.get(attr))
        has_fill = any(_local(k) in FILLS for k in rpr)
        if not has_fill:
            for k in d:
                if _local(k) in FILLS:
                    rpr.append(copy.deepcopy(k))
        if rpr.find(qn("a:latin")) is None and d.find(qn("a:latin")) is not None:
            rpr.append(copy.deepcopy(d.find(qn("a:latin"))))
        _sort_rpr(rpr)


NUM_RE = re.compile(r"^[\s\d,.%−\-–+<>≤≥и×]+$")


def is_kpi(text):
    t = text.strip()
    return bool(t) and bool(re.search(r"\d", t)) and bool(NUM_RE.match(t))


def iter_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # group
            yield from iter_shapes(sh.shapes)
        else:
            yield sh


def shape_by_name(slide, name):
    for sh in iter_shapes(slide.shapes):
        if sh.name == name:
            return sh
    raise KeyError(f"shape {name!r} not found")


def delete_shape(sh):
    el = sh._element
    el.getparent().remove(el)


def delete_by_names(slide, names):
    for n in names:
        try:
            delete_shape(shape_by_name(slide, n))
        except KeyError:
            print("  !! missing shape", n)


def style_title(sh, x=36, y=20, w=888, h=54, reposition=True, align="l"):
    tf = sh.text_frame
    body = tf._txBody
    bp = body.find(qn("a:bodyPr"))
    for k in ("lIns", "tIns", "rIns", "bIns"):
        bp.set(k, "0")
    bp.set("anchor", "t")
    bp.set("wrap", "square")
    for k in list(bp):
        if _local(k) in ("spAutoFit", "normAutofit", "noAutofit"):
            bp.remove(k)
    etree.SubElement(bp, qn("a:noAutofit"))
    for p in tf.paragraphs:
        pel = p._p
        materialize_defrpr(pel)
        ppr = pel.get_or_add_pPr()
        ppr.set("algn", align)
        for k in list(ppr):
            if _local(k) in ("lnSpc",):
                ppr.remove(k)
        ln = etree.Element(qn("a:lnSpc"))
        sp = etree.SubElement(ln, qn("a:spcPct"))
        sp.set("val", "95000")
        ppr.insert(0, ln)
        for r in pel.findall(qn("a:r")):
            style_rpr(get_rpr(r), size=20, font=TITLE_FONT, color=PURPLE)
        e = pel.find(qn("a:endParaRPr"))
        if e is not None:
            style_rpr(e, size=20, font=TITLE_FONT)
    if reposition:
        sh.left, sh.top, sh.width, sh.height = Pt(x), Pt(y), Pt(w), Pt(h)


def normalize_body(slide, skip_ids, kpi_size=28, lo=14, hi=18, font=BODY_FONT,
                   size_overrides=None, keep_big=()):
    """Clamp every run outside skip_ids to [lo, hi] pt and switch it to the body font.

    Numeric callouts (KPI) that were >= 22 pt are set to kpi_size.
    size_overrides: {shape_name: size} forces a size for all runs of that shape.
    keep_big: shape names whose text keeps its size (only font is changed).
    """
    size_overrides = size_overrides or {}
    report = []
    sp_tree = slide.shapes._spTree
    for el in sp_tree.iter(qn("p:sp"), qn("p:graphicFrame")):
        cnv = el.find(".//p:cNvPr", NS)
        name = cnv.get("name") if cnv is not None else ""
        if id(el) in skip_ids or el in skip_ids:
            continue
        for p_el in el.iter(qn("a:p")):
            materialize_defrpr(p_el)
            runs = p_el.findall(qn("a:r"))
            for r in runs:
                t = r.find(qn("a:t")).text or ""
                rpr = get_rpr(r)
                sz = effective_size(r)
                new = None
                if name in size_overrides:
                    new = size_overrides[name]
                elif name in keep_big:
                    new = None
                elif sz is None:
                    new = None
                elif is_kpi(t) and sz >= 22:
                    new = kpi_size
                elif sz > hi:
                    new = hi
                elif sz < lo:
                    new = lo
                final = new if new is not None else sz
                if final is not None and final != int(final) and name not in size_overrides:
                    new = max(lo, int(final))
                if new is not None and (sz is None or abs(new - sz) > 0.05):
                    report.append((name, sz, new, t[:30]))
                style_rpr(rpr, size=new, font=font, drop_links=True)
                # blue link colour -> black
                c = rpr.find("a:solidFill/a:srgbClr", NS)
                if c is not None and c.get("val") in ("467886", "0563C1", "0000FF") and name != "__chem__":
                    c.set("val", BLACK)
                if rpr.get("u") and c is not None and c.get("val") == BLACK:
                    del rpr.attrib["u"]
            e = p_el.find(qn("a:endParaRPr"))
            if e is not None:
                esz = e.get("sz")
                if esz:
                    v = int(esz) / 100
                    if name in size_overrides:
                        v = size_overrides[name]
                    else:
                        v = min(max(v, lo), hi)
                    e.set("sz", str(int(v * 100)))
    return report


# ---------- drawing helpers ----------

def _drop_style(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def find_text_shape(slide, needle):
    for sh in iter_shapes(slide.shapes):
        if sh.has_text_frame and needle in sh.text_frame.text:
            return sh
    raise KeyError(needle)


def replace_in_runs(shape, old, new):
    """Replace text inside a paragraph that may be split into several runs; keeps first run's format."""
    done = False
    for p in shape.text_frame.paragraphs:
        full = "".join(r.text for r in p.runs)
        if old in full:
            runs = list(p.runs)
            runs[0].text = full.replace(old, new)
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
            done = True
    if not done:
        raise KeyError(old)


def fix_chart(chart, lo=14, hi=18, font="Arial Narrow"):
    cs = chart._chartSpace
    for el in cs.iter():
        tag = etree.QName(el).localname
        if tag in ("defRPr", "rPr", "endParaRPr") and el.get("sz"):
            v = int(el.get("sz")) / 100
            el.set("sz", str(int(min(max(v, lo), hi) * 100)))
        if tag == "latin":
            el.set("typeface", font)
        if tag == "numFmt" and el.get("formatCode") in ("0.0%", "0%"):
            el.set("formatCode", "[$-419]" + el.get("formatCode"))
    chart.font.name = font  # default for text without its own typeface


def _fill(shape, color):
    if color is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor.from_string(color)


def _line(shape, color=None, width=None):
    if color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = RGBColor.from_string(color)
        shape.line.width = Pt(width or 1)


def add_box(slide, x, y, w, h, fill=GREY_T, line=None, line_w=1, radius=6, name=None, shape=None):
    kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
    sh = slide.shapes.add_shape(kind, Pt(x), Pt(y), Pt(w), Pt(h))
    _fill(sh, fill)
    _line(sh, line, line_w)
    sh.shadow.inherit = False
    _drop_style(sh)
    if kind == MSO_SHAPE.ROUNDED_RECTANGLE and radius:
        sh.adjustments[0] = min(0.5, radius / min(w, h))
    if name:
        sh.name = name
    sh.text_frame.text = ""
    return sh


def add_circle(slide, x, y, d, fill, line=None, name=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.OVAL, Pt(x), Pt(y), Pt(d), Pt(d))
    _fill(sh, fill)
    _line(sh, line)
    sh.shadow.inherit = False
    _drop_style(sh)
    if name:
        sh.name = name
    return sh


def add_line(slide, x1, y1, x2, y2, color=GREY, width=1.5, arrow=False, name=None):
    c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x1), Pt(y1), Pt(x2), Pt(y2))
    c.line.color.rgb = RGBColor.from_string(color)
    c.line.width = Pt(width)
    _drop_style(c)
    if arrow:
        ln = c.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        tail.set("w", "med")
        tail.set("len", "med")
    if name:
        c.name = name
    return c


def add_text(slide, x, y, w, h, paras, anchor="t", name=None, margin=0, font=BODY_FONT,
             shape=None, line_spacing=None, space_after=None):
    """paras: list of paragraphs; each paragraph is a list of runs or a dict.

    run = (text, size, bold, color) ; paragraph dict keys: runs, align, space_after.
    A bare string paragraph is allowed: ("text") -> 16 pt black.
    """
    if shape is None:
        tb = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(h))
    else:
        tb = shape
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    for side in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, side, Pt(margin))
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    first = True
    for para in paras:
        if isinstance(para, str):
            para = {"runs": [(para, 16, False, BLACK)]}
        elif isinstance(para, list):
            para = {"runs": para}
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[para.get("align", "l")]
        sa = para.get("space_after", space_after)
        if sa is not None:
            p.space_after = Pt(sa)
        sb = para.get("space_before")
        if sb is not None:
            p.space_before = Pt(sb)
        ls = para.get("line_spacing", line_spacing)
        if ls is not None:
            p.line_spacing = ls
        for (text, size, bold, color) in para["runs"]:
            r = p.add_run()
            r.text = text
            style_rpr(get_rpr(r._r), size=size, font=font, color=color, bold=bold)
    if name:
        tb.name = name
    return tb


def set_footer(slide, notes=(), sources=(), sources_ru=(), x=36, w=850, bottom=534, name="Источники"):
    """Single 8 pt footer: explanatory notes (RU) first, then sources (EN), then sources (RU)."""
    lines = [("n", t) for t in notes] + [("s", t) for t in sources] + [("r", t) for t in sources_ru]
    if not lines:
        return None
    # rough line estimate at 8 pt Calibri: ~215 characters per 888 pt
    per_line = int(215 * w / 888)
    n_lines = sum(max(1, -(-len(t) // per_line)) for _, t in lines)
    h = n_lines * 9.8 + 2
    paras = [{"runs": [(t, 8, False, BLACK)], "line_spacing": 1.0} for _, t in lines]
    return add_text(slide, x, bottom - h, w, h, paras, anchor="b", name=name)


def set_notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text


def white_bg(slide):
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(255, 255, 255)


def replace_everywhere(prs, old, new, slides=None):
    """Replace text across runs of every paragraph, keeping run formatting where possible."""
    count = 0
    if old == new or not old:
        return 0
    targets = list(prs.slides) if slides is None else slides
    for sl in targets:
        for p_el in sl.shapes._spTree.iter(qn("a:p")):
            start = 0
            while True:
                runs = p_el.findall(qn("a:r"))
                texts = [(r.find(qn("a:t")).text or "") for r in runs]
                full = "".join(texts)
                idx = full.find(old, start)
                if idx < 0:
                    break
                end = idx + len(old)
                pos = 0
                s_i = e_i = None
                for i, t in enumerate(texts):
                    if s_i is None and idx < pos + len(t):
                        s_i, s_off = i, idx - pos
                    if end <= pos + len(t):
                        e_i, e_off = i, end - pos
                        break
                    pos += len(t)
                ts = runs[s_i].find(qn("a:t"))
                if s_i == e_i:
                    ts.text = texts[s_i][:s_off] + new + texts[s_i][e_off:]
                else:
                    ts.text = texts[s_i][:s_off] + new
                    runs[e_i].find(qn("a:t")).text = texts[e_i][e_off:]
                    for r in runs[s_i + 1:e_i]:
                        p_el.remove(r)
                for r in p_el.findall(qn("a:r")):
                    t = r.find(qn("a:t"))
                    if t.text and (t.text != t.text.strip()):
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                start = idx + len(new)
                count += 1
    return count


def axis_percent_format(chart):
    cs = chart._chartSpace
    C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    for ax in cs.iter("{%s}valAx" % C):
        nf = ax.find("{%s}numFmt" % C)
        if nf is not None and "%" in nf.get("formatCode", ""):
            nf.set("formatCode", "0%")
            nf.set("sourceLinked", "0")


# ---------- reference (Levin) style blocks ----------
L_PURPLE, L_MAGENTA, L_BLUE, L_GREEN = "7030A0", "A02B93", "2C71A4", "4EA72E"
L_CYAN, L_TEAL, L_ORANGE, L_RED = "0F9ED5", "156082", "E97132", "B5323C"
L_CARD, L_CHEVRON, L_DARK = "F2F2F2", "D9D9D9", "404040"


def bar(slide, x, y, w, h, color, text, size=16, name=None, bold=False, align="l"):
    """Solid rounded bar with white text (reference style, slide 4 of the Levin deck)."""
    sh = add_box(slide, x, y, w, h, fill=color, radius=min(8, h / 4), name=name)
    tf = sh.text_frame
    tf.word_wrap = True
    tf.margin_left, tf.margin_right = Pt(10), Pt(8)
    tf.margin_top, tf.margin_bottom = Pt(2), Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER}[align]
    r = p.add_run()
    r.text = text
    style_rpr(get_rpr(r._r), size=size, font=BODY_FONT, color="FFFFFF", bold=bold)
    return sh


def card(slide, x, y, w, h, name=None, fill=None):
    return add_box(slide, x, y, w, h, fill=fill or L_CARD, radius=8, name=name)


def frame(slide, x, y, w, h, color, name=None, width=1.25):
    sh = add_box(slide, x, y, w, h, fill=None, line=color, line_w=width, radius=6, name=name)
    return sh


def arrow_down(slide, x, y, w=26, h=22, name=None, color=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, Pt(x), Pt(y), Pt(w), Pt(h))
    _fill(sh, color or L_CHEVRON)
    _line(sh, None)
    _drop_style(sh)
    if name:
        sh.name = name
    return sh


def chevron(slide, x, y, w=22, h=30, name=None, color=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Pt(x), Pt(y), Pt(w), Pt(h))
    _fill(sh, color or L_CHEVRON)
    _line(sh, None)
    _drop_style(sh)
    if name:
        sh.name = name
    return sh


def set_text_color(shape, color, size=None, bold=None):
    for r in shape._element.iter(qn("a:r")):
        style_rpr(get_rpr(r), color=color, size=size, bold=bold)


def anchor_middle(shape):
    shape.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE


def find_shapes(slide, name):
    return [sh for sh in iter_shapes(slide.shapes) if sh.name == name]


def set_connector(sh, begin=None, end=None):
    # drop glue to other shapes, otherwise viewers re-route the line back
    for tag in ("stCxn", "endCxn"):
        for el in sh._element.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}" + tag):
            el.getparent().remove(el)
    if begin:
        sh.begin_x, sh.begin_y = Pt(begin[0]), Pt(begin[1])
    if end:
        sh.end_x, sh.end_y = Pt(end[0]), Pt(end[1])


def nearest_connector(slide, name, pt_xy):
    best, bd = None, 1e9
    for sh in iter_shapes(slide.shapes):
        if sh.name == name and sh.shape_type == 9:
            for (x, y) in [(sh.begin_x, sh.begin_y), (sh.end_x, sh.end_y)]:
                d = (Emu(x).pt - pt_xy[0]) ** 2 + (Emu(y).pt - pt_xy[1]) ** 2
                if d < bd:
                    best, bd = sh, d
    return best


def label(slide, x, y, w, h, text, size=14, color="202020", bold=True, name=None, fill=None, align="l"):
    tb = add_text(slide, x, y, w, h, [{"runs": [(text, size, bold, color)], "align": align}], name=name)
    if fill:
        tb.fill.solid()
        tb.fill.fore_color.rgb = RGBColor.from_string(fill)
    return tb


def pointer(slide, x1, y1, x2, y2, color="303030", width=1.0, name=None):
    return add_line(slide, x1, y1, x2, y2, color=color, width=width, arrow=True, name=name)


def set_paragraph_texts(shape, lines):
    """Rewrite a text frame to exactly `lines` paragraphs, keeping the first run format of each paragraph."""
    tf = shape.text_frame
    paras = list(tf.paragraphs)
    # template paragraphs
    while len(paras) < len(lines):
        new_p = copy.deepcopy(paras[-1]._p)
        paras[-1]._p.addnext(new_p)
        paras = list(tf.paragraphs)
    for p in paras[len(lines):]:
        p._p.getparent().remove(p._p)
    for p, text in zip(tf.paragraphs, lines):
        runs = p._p.findall(qn("a:r"))
        for br in p._p.findall(qn("a:br")):
            p._p.remove(br)
        if not runs:
            r = p.add_run()
            runs = [r._r]
        runs[0].find(qn("a:t")).text = text
        for r in runs[1:]:
            p._p.remove(r)


def shape_lines(shape):
    out = []
    for p in shape.text_frame.paragraphs:
        out.append("".join(r.text for r in p.runs).replace("\x0b", "\n"))
    return "\n".join(out)


def replace_block(slide, old, new):
    """Replace text that may span paragraphs (old/new use ' / ' as a break). Returns number of shapes changed."""
    o = old.replace(" / ", "\n")
    n_ = new.replace(" / ", "\n")
    count = 0
    for sh in iter_shapes(slide.shapes):
        if not sh.has_text_frame:
            continue
        full = shape_lines(sh)
        if o in full:
            set_paragraph_texts(sh, full.replace(o, n_).split("\n"))
            count += 1
    return count


def set_rich_lines(shape, lines, color="000000"):
    """color=None keeps the template run colour."""
    """Rewrite a text frame as paragraphs of (text, bold) runs, keeping the first run's font and size."""
    tf = shape.text_frame
    tmpl = None
    for p in tf.paragraphs:
        rs = p._p.findall(qn("a:r"))
        if rs:
            tmpl = copy.deepcopy(rs[0].find(qn("a:rPr")))
            break
    set_paragraph_texts(shape, [""] * len(lines))
    for p, runs in zip(tf.paragraphs, lines):
        for r in p._p.findall(qn("a:r")):
            p._p.remove(r)
        for br in p._p.findall(qn("a:br")):
            p._p.remove(br)
        end = p._p.find(qn("a:endParaRPr"))
        for text, bold in runs:
            r = p.add_run()
            if tmpl is not None:
                old = r._r.find(qn("a:rPr"))
                if old is not None:
                    r._r.remove(old)
                r._r.insert(0, copy.deepcopy(tmpl))
            r.text = text
            r.font.bold = bool(bold)
            if color is not None:
                r.font.color.rgb = RGBColor.from_string(color)
            if end is not None:
                end.addprevious(r._r)
