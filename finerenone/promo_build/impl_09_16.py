"""Promo rework, slides 9-16 of base47.pptx (source of edits: promo/final_09_16.md).
Slide numbers are the numbers in base47.pptx. The order of slides, joints and the chevron chain are
handled by the caller. Every text edit raises KeyError if its target is not found."""
import copy

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import lib
import r14lib as L
from measure import n_lines, width

NB = " "
BLACK, WHITE = "000000", "FFFFFF"
CARD, CHEV, GREY, GREY_L = "F2F2F2", "D9D9D9", "7F7F7F", "AEAEAE"
TEAL, BLUE, ORANGE, GREEN_D, PURPLE = "156082", "2C71A4", "E97132", "196B24", "7030A0"
CYAN, LGREEN = "0F9ED5", "4EA72E"
FOOT_BOTTOM = 534.0


# ------------------------------------------------------------------ helpers
def nb(text):
    """Non-breaking space between a number and its unit (slide text only)."""
    import re
    t = re.sub(r"(?<=\d) (?=(?:мг/г|мг/ммоль|мг|ммоль/л|мл/мин|м²|недели|недель|месяца|месяцев|года|лет|типа|ПГ|ПП))",
               NB, text)
    return t


def geom(sh, x=None, y=None, w=None, h=None):
    if x is not None:
        sh.left = Pt(x)
    if y is not None:
        sh.top = Pt(y)
    if w is not None:
        sh.width = Pt(w)
    if h is not None:
        sh.height = Pt(h)
    return sh


def fill(sh, color):
    lib._fill(sh, color)
    return sh


def no_line(sh):
    sh.line.fill.background()
    return sh


def outline(sh, color, w=1.0):
    lib._line(sh, color, w)
    return sh


def round_rect(sh, radius=8):
    """Switch the geometry of an existing shape / text box to a rounded rectangle."""
    g = sh._element.spPr.find(qn("a:prstGeom"))
    g.set("prst", "roundRect")
    av = g.find(qn("a:avLst"))
    if av is None:
        av = etree.SubElement(g, qn("a:avLst"))
    for k in list(av):
        av.remove(k)
    gd = etree.SubElement(av, qn("a:gd"))
    gd.set("name", "adj")
    m = min(Emu(sh.width).pt, Emu(sh.height).pt)
    gd.set("fmla", "val %d" % int(min(50000, radius / m * 100000)))
    return sh


def insets(sh, l=0, t=0, r=0, b=0, anchor=None):
    tf = sh.text_frame
    tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = Pt(l), Pt(t), Pt(r), Pt(b)
    if anchor:
        tf.vertical_anchor = {"t": lib.MSO_ANCHOR.TOP, "m": lib.MSO_ANCHOR.MIDDLE,
                              "b": lib.MSO_ANCHOR.BOTTOM}[anchor]
    return sh


def clear_text(sh):
    """Remove all text of a shape but keep its first paragraph (and its paragraph properties)."""
    tb = sh.text_frame._txBody
    ps = tb.findall(qn("a:p"))
    for p in ps[1:]:
        tb.remove(p)
    for k in list(ps[0]):
        if k.tag in (qn("a:r"), qn("a:br"), qn("a:fld")):
            ps[0].remove(k)


def write(sh, paras, anchor="t", ml=0, mt=0, mr=0, mb=0):
    """Rewrite the text of a shape: paras as in lib.add_text (runs = (text, size, bold, color))."""
    clear_text(sh)
    lib.add_text(None, 0, 0, 0, 0, paras, anchor=anchor, shape=sh, margin=0)
    insets(sh, ml, mt, mr, mb)
    return sh


def move_before(sh, ref):
    ref._element.addprevious(sh._element)
    return sh


def recolor(sh, color, text=False):
    """Set every srgbClr inside the spPr (fill and line) of a shape to `color`."""
    for c in sh._element.spPr.iter(qn("a:srgbClr")):
        c.set("val", color)


def line_color(sh, color, width=None):
    """Colour (and optionally width) of a connector / line shape; arrowheads stay."""
    sh.line.color.rgb = RGBColor.from_string(color)
    if width:
        sh.line.width = Pt(width)


def solid_fill_of(sh, color, line=None, line_w=None):
    """Replace a gradient / any fill of an existing shape by a solid fill; optionally recolour the outline."""
    spPr = sh._element.spPr
    for k in list(spPr):
        if etree.QName(k).localname in ("solidFill", "gradFill", "noFill", "pattFill", "blipFill"):
            spPr.remove(k)
    sf = etree.Element(qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr")).set("val", color)
    g = spPr.find(qn("a:prstGeom"))
    if g is None:
        g = spPr.find(qn("a:custGeom"))
    g.addnext(sf)
    if line:
        ln = spPr.find(qn("a:ln"))
        if ln is not None:
            for c in ln.iter(qn("a:srgbClr")):
                c.set("val", line)
            if line_w:
                ln.set("w", str(int(line_w * 12700)))


def shift_y(shapes, dy):
    for sh in shapes:
        sh.top = sh.top + Pt(dy)


def drop(s, *names):
    """Delete shapes by name (every shape with that name)."""
    for n in names:
        found = lib.find_shapes(s, n)
        if not found:
            raise KeyError(n)
        for sh in found:
            lib.delete_shape(sh)


def line_h(size):
    return size * 1.2


def _stack_shape(s, name, text, size):
    try:
        sh = L.shape_named(s, name)
        if text is not None:
            L.set_lines(sh, [text])
    except KeyError:
        if text is None:
            raise
        sh = lib.add_text(s, 36, 480, 850, 20, [[(text, size, False, BLACK)]], name=name)
    lib.set_text_color(sh, BLACK, size=size)
    return sh


def footers(s, abbr=None, src=None, star=None, width=850.0, bottom=FOOT_BOTTOM):
    """Stack the footnotes from the bottom edge up: sources (8 pt) lowest, abbreviations (10 pt) above,
    the asterisk line (8 pt) topmost. A text left as None keeps what the shape has now. Returns the top y."""
    items = []
    if src is not None or lib.find_shapes(s, "Источники"):
        items.append(("Источники", src, 8))
    if abbr is not None or lib.find_shapes(s, "Сокращения"):
        items.append(("Сокращения", abbr, 10))
    if star is not None:
        items.append(("Сноска", star, 8))
    y = bottom
    for name, text, size in items:
        sh = _stack_shape(s, name, text, size)
        t = "".join(p.text for p in sh.text_frame.paragraphs)
        n = n_lines([(t, False)], width, size)
        h = n * line_h(size) + 1
        y -= h
        geom(sh, 36, y, width, h)
        insets(sh, 0, 0, 0, 0, anchor="b")
        sh.text_frame.word_wrap = True
    return y


# ---- notes
def notes_blocks(s):
    return [p.text for p in s.notes_slide.notes_text_frame.paragraphs if p.text.strip()]


def notes_write(s, blocks):
    """Write notes as paragraphs separated by an empty paragraph (as in the original notes)."""
    L.set_notes(s, "\n\n".join(blocks))


def notes_add(s, *blocks):
    notes_write(s, notes_blocks(s) + list(blocks))


def notes_prepend(s, text):
    b = notes_blocks(s)
    b[0] = text + " " + b[0]
    notes_write(s, b)


def visit(text):
    return "Фраза для визита: «" + text + "»"


def ask(question, answer):
    return "Если врач спросит: «" + question + "» Ответ: «" + answer + "»"


# ------------------------------------------------------------------ slide 9
def s09(prs):
    s = L.S(prs, 9)
    L.set_title(s, "При ХБП растёт нагрузка на сердце: защита нужна и почкам, и сердцу")

    # subtitle: Bold 16 pt black, no type number; the long context line goes (its sense moves to the notes)
    sub = L.shape_named(s, "classification_2")
    L.set_lines(sub, ["Ренокардиальный синдром: болезнь почек ухудшает работу сердца"])
    lib.set_text_color(sub, BLACK, size=16, bold=True)
    drop(s, "clinical_context_2")

    # the scheme (kidney, three rows, heart) moves up by 30 pt
    for n in ("kidney_source_label", "heart_target_label", "heart_changes", "Почка", "Сердце",
              "pathway_4_1", "pathway_4_2", "pathway_4_3"):
        shift_y([L.shape_named(s, n)], -30)
    drop(s, "Connector 9", "Connector 10", "Connector 14", "Connector 15", "Connector 19")  # blue bars, rules
    for n in (7, 11, 12, 16, 17, 20, 21, 22, 23, 24):
        c = L.shape_named(s, "Connector %d" % n)
        c.top = c.top + Pt(-30)
        line_color(c, GREY)
    # arrows into the rows now end at the cards (x 172) and the right ticks start at their edge (x 634)
    for n in (11, 16, 20):
        c = L.shape_named(s, "Connector %d" % n)
        c.width = Pt(170 - Emu(c.left).pt)
    for n in (12, 17, 21):
        c = L.shape_named(s, "Connector %d" % n)
        right = Emu(c.left).pt + Emu(c.width).pt
        c.left = Pt(636)
        c.width = Pt(right - 636)
    lib.set_text_color(L.shape_named(s, "heart_target_label"), BLACK)
    lib.set_text_color(L.shape_named(s, "kidney_source_label"), BLACK)

    # three rows: grey cards without outline, 16 pt black, «альдостерона» Bold in the first one
    rows = [("pathway_4_1", 126, 55), ("pathway_4_2", 193, 54), ("pathway_4_3", 259, 42)]
    for name, y, h in rows:
        sh = L.shape_named(s, name)
        geom(sh, 172, y, 462, h)
        fill(sh, CARD)
        no_line(sh)
        round_rect(sh, 8)
    write(L.shape_named(s, "pathway_4_1"),
          [[("Активация ренин-ангиотензин-альдостероновой системы, избыток ", 16, False, BLACK),
            ("альдостерона", 16, True, BLACK),
            (". Задержка натрия и воды. Увеличение объёма жидкости.", 16, False, BLACK)]],
          anchor="m", ml=6, mr=6, mt=2, mb=2)
    write(L.shape_named(s, "pathway_4_2"),
          [[("Активация симпатической нервной системы. Сужение сосудов. Повышение уровня катехоламинов, "
             "прежде всего норадреналина.", 16, False, BLACK)]], anchor="m", ml=6, mr=6, mt=2, mb=2)
    write(L.shape_named(s, "pathway_4_3"),
          [[("Уремия, электролитные нарушения и снижение pH крови (ацидемия).", 16, False, BLACK)]],
          anchor="m", ml=6, mr=6, mt=2, mb=2)

    # shared processes: teal plate + one grey card with four columns separated by white 1 pt rules
    drop(s, "Connector 29", "Connector 31", "Connector 33", "Connector 35", "Connector 37", "Connector 38")
    plate = L.shape_named(s, "shared_title")
    geom(plate, 36, 326, 850, 26)
    fill(plate, TEAL)
    no_line(plate)
    round_rect(plate, 6)
    write(plate, [[("Общие для обоих органов процессы: окислительный стресс и воспаление", 16, True, WHITE)]],
          anchor="m", ml=12, mr=8)
    card_y, card_h = 354, 66
    card = lib.card(s, 36, card_y, 850, card_h, name="Карточка_общие_процессы")
    cols = [("shared_0.57", "Больше воспалительных веществ (цитокинов)"),
            ("shared_3.75", "Нарушение баланса оксида азота и активных форм кислорода"),
            ("shared_6.93", "Нарушение функции митохондрий"),
            ("shared_10.05", "Нарушение работы эндотелия (внутренней стенки сосудов)")]
    first = L.shape_named(s, cols[0][0])
    move_before(card, first)
    cw = 850 / 4
    for i, (name, text) in enumerate(cols):
        sh = L.shape_named(s, name)
        geom(sh, 36 + i * cw, card_y, cw, card_h)
        write(sh, [{"runs": [(text, 16, False, BLACK)], "align": "c"}], anchor="m", ml=10, mr=10)
    for i in (1, 2, 3):
        x = 36 + i * cw
        d = lib.add_line(s, x, card_y + 8, x, card_y + card_h - 8, color=WHITE, width=1.0,
                         name="Разделитель_колонок_%d" % i)
        card._element.addnext(d._element)

    # bottom purple plate: only the finerenone message
    link = L.shape_named(s, "Связь_с_финереноном")
    geom(link, 36, 436, 850, 38)
    fill(link, PURPLE)
    no_line(link)
    round_rect(link, 8)
    write(link, [[("Избыток альдостерона действует на сердце и почки через минералокортикоидный рецептор "
                   "(МКР). ", 16, False, WHITE),
                  ("Финеренон блокирует МКР.", 16, True, WHITE)]], anchor="m", ml=12, mr=8)

    footers(s, abbr="ХБП: хроническая болезнь почек; МКР: минералокортикоидный рецептор.",
            src="Kumar U, et al. Cardiol Clin. 2019;37:251–265. Ndumele CE, et al. Circulation. "
                "2026;154:e50–e158. Agarwal R, et al. Eur Heart J. 2021;42:152–161. "
                "Адаптировано по: Servier Medical Art (CC BY 4.0).")

    # ---- notes
    L.rep(s, "Нарушения состава крови и уремия дополнительно влияют на работу сердца.",
          "Нарушения состава крови и уремия дополнительно влияют на работу сердца. Такое сочетание называют "
          "хроническим ренокардиальным синдромом, тип 4: болезнь почек ухудшает работу сердца.", notes=True)
    L.rep(s, "Эти механизмы помогают понять заболевание, но не заменяют данные исследований конкретного "
             "препарата.",
          "Руководство AHA/ACC/ADA/ASN 2026 рассматривает болезни сердца, почек и обмен веществ как "
          "взаимосвязанные состояния. Доказательства для финеренона дают клинические исследования, их "
          "разберём дальше.", notes=True)
    L.rep(s, "При обсуждении финеренона ориентируются на группу пациентов, указанную в местной инструкции: "
             "взрослые с ХБП, связанной с диабетом 2 типа. Сопутствующая сердечная недостаточность сама по "
             "себе не исключает эту группу. Отдельное показание при сердечной недостаточности в "
             "предоставленной инструкции не указано.",
          "Один из путей на схеме идёт через избыток альдостерона. Он действует на сердце и почки через МКР, "
          "а финеренон блокирует этот рецептор. Финеренон назначают взрослым с ХБП, связанной с диабетом "
          "2 типа: показание включает снижение риска почечных и сердечно-сосудистых событий, в том числе "
          "госпитализации по поводу сердечной недостаточности. В объединённом анализе FIDELITY (13 026 "
          "пациентов) такой риск был ниже на 22% по сравнению с плацебо (ОР 0,78; 95% ДИ 0,66–0,92; медиана "
          "3,0 года). Теперь о том, как найти таких пациентов.", notes=True)
    notes_add(
        s,
        visit("Доктор, при диабете и болезни почек страдают и почки, и сердце. Финеренон блокирует рецептор "
              "альдостерона и снижает риск почечных и сердечно-сосудистых событий."),
        ask("А при сердечной недостаточности без хронической болезни почек?",
            "Показание относится к взрослым с ХБП, связанной с диабетом 2 типа. По другим группам пациентов "
            "передам ваш вопрос в медицинский отдел."),
        ask("У пациента с ХБП и диабетом есть и сердечная недостаточность. Можно ли назначать финеренон?",
            "Сердечная недостаточность в перечне противопоказаний инструкции не стоит и группу сама по себе "
            "не исключает. В FIDELIO-DKD и FIGARO-DKD не включали пациентов с хронической симптомной "
            "сердечной недостаточностью со сниженной фракцией выброса; сердечная недостаточность в анамнезе "
            "была у 7,7% участников FIDELITY. Решение о назначении принимает врач."))


# ------------------------------------------------------------------ slide 10
def s10(prs):
    s = L.S(prs, 10)
    L.set_title(s, nb("При диабете 2 типа рСКФ и АКО проверяют не реже раза в год: АКО выявляет болезнь почек "
                      "при сохранной фильтрации"))
    # question plate: violet is for finerenone only
    fill(L.shape_named(s, "Плашка_вопрос"), BLUE)
    lib.set_text_color(L.shape_named(s, "Clinical_question"), WHITE, size=18, bold=True)

    # right column: three grey cards A (answer + number), B (frequency), C (question to the doctor)
    a = L.shape_named(s, "Карточка_ответ")
    geom(a, 612, 156, 312, 186)
    cb = lib.card(s, 612, 350, 312, 58, name="Карточка_частота")
    cc = lib.card(s, 612, 416, 312, 76, name="Карточка_вопрос_врачу")
    geom(L.shape_named(s, "Карточка_моча"), h=166)
    a._element.addnext(cb._element)
    cb._element.addnext(cc._element)

    ans = L.shape_named(s, "Answer")
    geom(ans, 628, 164, 282, 66)
    num = lib.add_text(s, 628, 234, 282, 36, [[("40%", 28, True, BLUE)]], anchor="m", name="Число_40")
    cap = lib.add_text(s, 628, 272, 282, 62,
                       [[("участников FIDELITY имели рСКФ 60 и выше. ХБП у них определяли только по "
                          "альбуминурии.", 16, False, BLACK)]], name="Подпись_40")
    fr = L.shape_named(s, "Frequency")
    geom(fr, 628, 350, 282, 58)
    write(fr, [[("При ХБП", 16, True, BLACK),
                (" частоту контроля определяет категория риска: чем выше риск, тем чаще.", 16, False, BLACK)]],
          anchor="m")
    dq = L.shape_named(s, "Dipstick_limitation")
    geom(dq, 628, 416, 282, 76)
    write(dq, [[("Вопрос врачу:", 16, True, BLACK),
                (nb(" у скольких ваших пациентов с диабетом 2 типа за год определяли и рСКФ, и АКО?"),
                 16, False, BLACK)]], anchor="m")
    for sh in (num, cap):
        move_before(sh, fr)

    footers(s,
            abbr="рСКФ: расчётная скорость клубочковой фильтрации; АКО: отношение альбумина к креатинину мочи; "
                 "ХБП: хроническая болезнь почек; FIDELITY: объединённый анализ FIDELIO-DKD и FIGARO-DKD, "
                 "13" + NB + "026 пациентов с ХБП и диабетом 2" + NB + "типа; рСКФ 60 и выше у 5195 (39,9%).",
            src="KDIGO 2024. Kidney Int. 105(4S):S117–S314. American Diabetes Association Professional "
                "Practice Committee. Standards of Care in Diabetes-2026. Diabetes Care. 2026;49(Suppl 1):"
                "S246–S260. Agarwal R, et al. Eur Heart J. 2022;43:474–484.")

    # ---- notes
    notes_prepend(s, "Искать начинаем с анализов.")
    L.rep(s, "При установленной хронической болезни почек частота зависит от стадии.",
          "При установленной хронической болезни почек частота зависит от категории риска.", notes=True)
    L.rep(s, "Повреждение возможно уже при сохранной фильтрации, поэтому анализ крови не заменяет определение "
             "АКО. Тест-полоска на общий белок также не заменяет этот анализ.",
          "Повреждение возможно уже при сохранной фильтрации. В FIDELITY у 40% участников (5195 из 13 026) "
          "рСКФ была 60 и выше, и ХБП у них определяли только по альбуминурии. Поэтому одной рСКФ "
          "недостаточно, нужно и АКО. Если альбуминурия выявлена, дальше нужно подтвердить ХБП. Критерии на "
          "следующем слайде.", notes=True)
    b = notes_blocks(s)
    quote = b.pop()          # the closing paragraph in guillemets was a phrase for the visit without a label
    assert quote.startswith("«При диабете 2 типа"), quote[:30]
    b.append(visit("Доктор, при диабете 2 типа рСКФ и АКО проверяют не реже раза в год. Повышенное АКО может "
                   "указать на болезнь почек даже при сохранной рСКФ. У скольких ваших пациентов за последний "
                   "год определяли оба показателя?"))
    b.append(ask("А если лаборатория выдаёт АКО в мг/ммоль?",
                 "30 мг/г соответствует примерно 3 мг/ммоль (KDIGO 2024)."))
    b.append(ask("Так бывает у 40% всех пациентов с ХБП и диабетом?",
                 "Нет. 40% относится к участникам FIDELITY: в исследования включали пациентов с "
                 "альбуминурией, в том числе с рСКФ 60 и выше. Частоту в практике число не показывает. Оно "
                 "показывает, что без АКО такие пациенты остаются вне критериев ХБП."))
    b.append(ask("Хватит ли тест-полоски?",
                 "Для скрининга альбуминурии определяют АКО. Вопрос о замене АКО другими тестами передам в "
                 "медицинский отдел."))
    notes_write(s, b)


# ------------------------------------------------------------------ slide 11
def s11(prs):
    s = L.S(prs, 11)
    L.set_title(s, nb("ХБП подтверждают по рСКФ ниже 60 или АКО от 30 мг/г, если это сохраняется не менее "
                      "3 месяцев"))
    L.rep(s, "Критерий ХБП: ≥30 мг/г", nb("Критерий ХБП: ≥30 мг/г (≥3 мг/ммоль)"), count=1)
    L.rep(s, "или маркера повреждения почек.", "или признака повреждения почек, например альбуминурии.",
          count=1)
    L.rep(s, "Сохранная рСКФ не исключает болезнь почек. ", "", count=1)
    # plates and thresholds: the colour is carried by the plates, thresholds are black Bold 18 pt
    fill(L.shape_named(s, "Плашка_повреждение"), ORANGE)
    for n in ("Плашка_функция", "Плашка_повреждение"):
        lib.set_text_color(L.shape_named(s, n), WHITE, size=18, bold=True)
    for n in ("function-threshold", "damage-threshold"):
        lib.set_text_color(L.shape_named(s, n), BLACK, size=18, bold=True)
    geom(L.shape_named(s, "Карточка_хроническая"), h=92)
    geom(L.shape_named(s, "transient"), h=24)
    footers(s)
    # ---- notes
    L.rep(s, "Хроническую болезнь почек можно выявить по снижению функции или признакам повреждения.",
          "Какие значения этих анализов говорят о ХБП? Хроническую болезнь почек можно выявить по снижению "
          "функции или признакам повреждения.", notes=True)
    L.rep(s, "Повышенный АКО подтверждают повторно", "Повышенное АКО подтверждают повторно", notes=True)
    L.rep(s, "Диагноз устанавливает врач.", "Диагноз устанавливает врач. Теперь о том, как два показателя "
                                            "вместе описывают риск.", notes=True)
    notes_add(s, visit("Доктор, по критериям KDIGO для ХБП достаточно одного признака. Если АКО от 30 мг/г "
                       "сохраняется три месяца, это ХБП и при сохранной рСКФ. Диагноз, конечно, ставите вы."))


# ------------------------------------------------------------------ slide 12
def s12(prs):
    s = L.S(prs, 12)
    L.set_title(s, nb("Альбуминурия повышает риск и при сохранной рСКФ: при рСКФ 65 и АКО 350 мг/г риск "
                      "высокий"))
    # table (header, row heads, 18 coloured cells, frames) moves down by 20 pt
    table = [sh for sh in s.shapes if Emu(sh.left).pt >= 230 and 80 < Emu(sh.top).pt < 450]
    assert len(table) == 5 + 6 + 5 + 18 + 18 - 5 - 6 + 2 or len(table) > 40, len(table)
    shift_y(table, 20)
    L.set_lines(L.shape_named(s, "Text_4"), ["рСКФ (G)", nb("мл/мин/1,73 м²")])
    for n in ("Рамка_пример_A1", "Рамка_пример_A3"):
        outline(L.shape_named(s, n), BLACK, 2.25)

    # left card: colour legend + example, 16 pt
    card = L.shape_named(s, "Карточка_пример")
    geom(card, 36, 181, 188, 206)
    ex = L.shape_named(s, "Text_50")
    geom(ex, 44, 189, 172, 190)
    write(ex, [{"runs": [("Цвет клетки: риск прогрессирования ХБП и осложнений в группе пациентов.", 16,
                          False, BLACK)], "space_after": 6},
               [("Пример:", 16, True, BLACK), (nb(" рСКФ 65 (G2)."), 16, False, BLACK)],
               [(nb("АКО 15 мг/г (A1): риск низкий*."), 16, False, BLACK)],
               [(nb("АКО 350 мг/г (A3): риск высокий."), 16, False, BLACK)]], anchor="m")

    # bracket over A2 and A3 with the label: who was included
    lab = lib.add_text(s, 546, 78, 378, 19,
                       [{"runs": [("Участники FIDELIO-DKD и FIGARO-DKD: АКО от 30" + NB + "мг/г", 14, False,
                                   BLACK)], "align": "c"}], name="Подпись_скобки")
    lib.add_line(s, 548, 101, 922, 101, color=BLUE, width=1.5, name="Скобка_A2_A3")
    lib.add_line(s, 548, 101, 548, 105, color=BLUE, width=1.5, name="Скобка_A2_A3_левый_край")
    lib.add_line(s, 922, 101, 922, 105, color=BLUE, width=1.5, name="Скобка_A2_A3_правый_край")

    footers(s, star="* При G1–G2 без альбуминурии и других признаков повреждения диагноз болезни почек не "
                    "устанавливают.",
            abbr="рСКФ: расчётная скорость клубочковой фильтрации; АКО: отношение альбумина к креатинину "
                 "мочи; G: категории рСКФ; A: категории АКО. АКО в мг/ммоль: A1 <3, A2 3–30, A3 >30.",
            src="KDIGO 2024. Kidney Int. 105(4S):S117–S314.")
    # ---- notes
    notes_prepend(s, "Здесь два показателя сведены в одну таблицу.")
    L.rep(s, "сама по себе не означает хроническую болезнь почек.",
          "сама по себе не означает хроническую болезнь почек. Дальше соберём портрет такого пациента.",
          notes=True)
    notes_add(s, visit("Доктор, при диабете АКО от 30 мг/г повышает риск даже при сохранной рСКФ. Поэтому "
                       "смотрим на оба показателя."))


# ------------------------------------------------------------------ slide 13
BLOCKS = [  # (frame, disc, number, title, left text, right text, colour, top)
    ("Рамка_группа_1", "Google Shape;63;p13", "Google Shape;64;p13", "Google Shape;65;p13",
     "Google Shape;66;p13", "Google Shape;67;p13", TEAL, 59),
    ("Рамка_группа_2", "Google Shape;71;p13", "Google Shape;72;p13", "Google Shape;73;p13",
     "Google Shape;74;p13", "Google Shape;75;p13", GREEN_D, 160),
    ("Рамка_группа_3", "Google Shape;79;p13", "Google Shape;80;p13", "Google Shape;81;p13",
     "Google Shape;82;p13", "Google Shape;83;p13", PURPLE, 261),
    ("Рамка_группа_4", "Google Shape;87;p13", "Google Shape;88;p13", "Google Shape;89;p13",
     "Google Shape;90;p13", "Google Shape;91;p13", ORANGE, 362),
]


def s13(prs):
    s = L.S(prs, 13)
    L.set_title(s, "Четыре группы препаратов защищают почки разными путями: у финеренона своя мишень, МКР")

    # text edits
    L.rep(s, "сердце, сосудах.", "сердце и сосудах.", count=1)
    L.rep(s, "блокирует его активацию.", "блокирует МКР, препятствуя включению генов воспаления и фиброза.",
          count=1)
    L.rep(s, "и сердечно-сосудистых событий.",
          nb("и сердечно-сосудистых событий у взрослых с ХБП и диабетом 2 типа (FIDELIO-DKD, FIGARO-DKD, "
             "FIDELITY)."), count=1)
    L.rep(s, "Польза семаглутида:", "Польза¹:", count=1)
    L.rep(s, " ниже риск прогрессирования ХБП или сердечно-сосудистой смерти.",
          " ниже риск тяжёлых почечных событий и смерти от почечных или сердечно-сосудистых причин.", count=1)
    L.rep(s, "Врач выбирает лечение по рСКФ, альбуминурии и калию; контролирует давление и уровень глюкозы в "
             "крови.",
          "Финеренон добавляют к ингибитору АПФ или БРА. Набор препаратов врач подбирает по рСКФ, АКО и калию.",
          count=1)
    L.set_lines(L.shape_named(s, "Google Shape;93;p13"), ["Где находится мишень"])

    # four blocks: coloured plate over a grey card, no outline; block 3 (finerenone) gets a violet frame
    for frame, disc, num, title, left, right, col, top in BLOCKS:
        f = L.shape_named(s, frame)
        geom(f, 244, top + 26, 680, 72)
        fill(f, CARD)
        no_line(f)
        if col == PURPLE:
            outline(f, PURPLE, 1.5)
        t = L.shape_named(s, title)
        geom(t, 244, top, 680, 24)
        fill(t, col)
        no_line(t)
        round_rect(t, 6)
        lib.set_text_color(t, WHITE, size=16, bold=True)
        insets(t, 34, 0, 8, 0, anchor="m")
        d = L.shape_named(s, disc)
        move_before(t, d)                      # the plate goes under the number disc
        geom(d, 248, top + 1, 22, 22)
        fill(d, WHITE)
        n_ = L.shape_named(s, num)
        geom(n_, 248, top + 1, 22, 22)
        lib.set_text_color(n_, col, size=14, bold=True)
        for nm, x, w in ((left, 254, 292), (right, 564, 352)):
            sh = L.shape_named(s, nm)
            geom(sh, x, top + 30, w, 64)
            insets(sh, 0, 0, 0, 0, anchor="t")
            lib.set_text_color(sh, BLACK, size=14)
    # the number discs on the drawing take the colour of the plates (block 4 was cyan)
    for n in ("membrane_target_4_disc",):
        fill(L.shape_named(s, n), ORANGE)
    bottom = L.shape_named(s, "Google Shape;95;p13")
    geom(bottom, 36, 462, 888, 24)

    footers(s,
            abbr="ХБП: хроническая болезнь почек; рСКФ: расчётная скорость клубочковой фильтрации; АКО: "
                 "отношение альбумина к креатинину мочи; АПФ: ангиотензинпревращающий фермент; БРА: блокатор "
                 "рецепторов ангиотензина II; НГЛТ2: натрий-глюкозный котранспортёр 2 типа; МКР: "
                 "минералокортикоидный рецептор; ГПП-1: глюкагоноподобный пептид-1.",
            src="¹ Семаглутид, исследование FLOW: Perkovic V, et al. N Engl J Med. 2024;391:109–121. KDIGO "
                "2024. Kidney Int. 105(4S):S117–S314. Финеренон-Орвилле. Инструкция по медицинскому "
                "применению. Bakris GL, et al. N Engl J Med. 2020;383:2219–2229. Pitt B, et al. N Engl J "
                "Med. 2021;385:2252–2263. Адаптировано по: Servier Medical Art (CC BY 4.0).")

    # ---- notes
    L.rep(s, "Для защиты почек при диабете 2 типа используют группы препаратов с разными мишенями.",
          "Пациента с ХБП и диабетом мы нашли. Чем его лечат сегодня? Для защиты почек при диабете 2 типа "
          "используют четыре группы препаратов с разными мишенями.", notes=True)
    L.rep(s, "В исследовании FLOW для семаглутида показана польза по комбинированному исходу, включавшему "
             "прогрессирование болезни почек или сердечно-сосудистую смерть. Данные конкретного препарата "
             "нельзя автоматически переносить на весь класс или на все сердечно-сосудистые события.",
          "В исследовании FLOW для семаглутида показана польза по комбинированному исходу: тяжёлые почечные "
          "события или смерть от почечных или сердечно-сосудистых причин.", notes=True)
    L.rep(s, "Схема не означает, что всем пациентам нужен одинаковый набор из четырёх групп. Давление и "
             "уровень глюкозы продолжают контролировать.",
          "Давление и уровень глюкозы продолжают контролировать. Финеренон добавляют к ингибитору АПФ или "
          "БРА. С остальными группами его сочетание разберём отдельно. Хватает ли этих препаратов? Ответ "
          "даёт исследование CREDENCE.", notes=True)
    notes_add(
        s,
        visit("Доктор, у финеренона своя мишень: минералокортикоидный рецептор. Его добавляют к ингибитору "
              "АПФ или БРА."),
        "Если врач спросит о доказательствах для ингибиторов АПФ и БРА: «Исходы при диабете 2 типа и "
        "нефропатии показаны для блокаторов рецепторов ангиотензина: лозартан в RENAAL (1513 пациентов; "
        "первичный комбинированный исход, то есть удвоение креатинина, терминальная почечная "
        "недостаточность или смерть, ниже на 16% по сравнению с плацебо, в среднем 3,4 года) и ирбесартан в "
        "IDNT (1715 пациентов с гипертензией; ниже на 20% по сравнению с плацебо, в среднем 2,6 года).»",
        ask("А других агонистов ГПП-1?",
            "Данные FLOW относятся к семаглутиду, по другим препаратам передам вопрос в медицинский отдел."),
        "Если врач спросит, нужны ли все четыре группы каждому: «Набор подбирает врач по рСКФ, АКО, калию, "
        "давлению и глюкозе; финеренон добавляют к ингибитору АПФ или БРА у взрослых с ХБП и диабетом 2 "
        "типа.»",
        "Если врач спросит о сочетании с ингибиторами НГЛТ2 и агонистами ГПП-1: «В FIDELITY ингибиторы НГЛТ2 "
        "в начале исследования принимали 6,7% участников (877 человек), агонисты ГПП-1 7,2% (944 человека); "
        "результаты в этих подгруппах согласуются с общими. Рандомизированного сравнения комбинаций в "
        "FIDELITY не было; сочетание с эмпаглифлозином изучено в CONFIDENCE, оно разобрано на слайде о "
        "комбинации.»")


# ------------------------------------------------------------------ slide 14
def s14(prs):
    s = L.S(prs, 14)
    L.set_title(s, "Отношение рисков 0,82 означает риск ниже на 18% по сравнению с плацебо")

    # banner: the study and the group (violet: about finerenone)
    ban = L.shape_named(s, "Пример исследования")
    geom(ban, 36, 84, 888, 34)
    fill(ban, PURPLE)
    no_line(ban)
    round_rect(ban, 8)
    write(ban, [[(nb("FIDELIO-DKD: взрослые с ХБП и диабетом 2 типа, финеренон или плацебо на фоне ингибитора "
                     "АПФ или БРА"), 18, True, WHITE)]], anchor="m", ml=12, mr=8)
    drop(s, "Контекст примера")

    # four grey cards: label on the left (Bold 16 pt), value on the right (18 pt)
    ys = [128, 190, 252, 314]
    first = L.shape_named(s, "Первичный исход")
    for i, y in enumerate(ys):
        c = lib.card(s, 36, y, 888, 56, name="Карточка_%d" % (i + 1))
        move_before(c, first)
    rows = [("Первичный исход", "Состав исхода"), ("Отношение рисков", "Эффект"),
            ("Доверительный интервал", "Неопределённость")]
    labels = ["Что оценивали", "Отношение рисков", "95% доверительный интервал"]
    for (ln, vn), y, text in zip(rows, ys, labels):
        lab = L.shape_named(s, ln)
        geom(lab, 48, y, 236, 56)
        write(lab, [[(text, 16, True, BLACK)]], anchor="m")
        val = L.shape_named(s, vn)
        geom(val, 292, y, 620, 56)
    write(L.shape_named(s, "Состав исхода"),
          [[("Почечная недостаточность, устойчивое снижение рСКФ на ≥40% или смерть от почечных причин.", 18,
             False, BLACK)]], anchor="m")
    write(L.shape_named(s, "Эффект"),
          [[("0,82", 28, True, PURPLE), (": риск ниже на 18% по сравнению с плацебо", 18, True, PURPLE)]],
          anchor="m")
    write(L.shape_named(s, "Неопределённость"),
          [[("0,73–0,93. Весь интервал ниже 1: результат статистически значим.", 18, False, BLACK)]],
          anchor="m")
    lab4 = lib.add_text(s, 48, ys[3], 236, 56, [[("В процентах и ЧБНЛ", 16, True, BLACK)]], anchor="m",
                        name="Метка_в_процентах_и_ЧБНЛ")
    val4 = lib.add_text(s, 292, ys[3], 620, 56,
                        [[(nb("17,8% на финереноне и 21,1% на плацебо за медиану 2,6 года"), 18, False, BLACK)],
                         [(nb("ЧБНЛ 29 за 3 года: столько пациентов нужно лечить, чтобы предотвратить одно "
                              "событие"), 18, False, BLACK)]], anchor="m", name="Значение_в_процентах_и_ЧБНЛ")
    # bottom line on a grey card
    cond = L.shape_named(s, "Условия интерпретации")
    geom(cond, 36, 392, 888, 48)
    fill(cond, CARD)
    no_line(cond)
    round_rect(cond, 8)
    write(cond, [[("Результат относится к комбинированному исходу у взрослых с ХБП и диабетом 2 типа на фоне "
                   "ингибитора АПФ или БРА.", 17, False, BLACK)]], anchor="m", ml=12, mr=12)

    footers(s,
            abbr="ХБП: хроническая болезнь почек; рСКФ: расчётная скорость клубочковой фильтрации; АПФ: "
                 "ангиотензинпревращающий фермент; БРА: блокатор рецепторов ангиотензина II; ОР: отношение "
                 "рисков; ЧБНЛ: число больных, которых необходимо лечить.")
    # ---- notes (the order of slides is fixed by the caller: after the slide on the way from mechanism to indication)
    notes_prepend(s, "Мы показали, откуда берётся показание. Теперь доказательства, и сначала короткая "
                     "памятка: как читать числа исследования и как пересказать их врачу.")
    L.rep(s, "Оно не показывает абсолютное число предотвращённых событий и не гарантирует тот же результат "
             "каждому пациенту.",
          "Абсолютную разницу видно в процентах: 21,1% против 17,8%, то есть на 3,3 процентного пункта "
          "меньше.", notes=True)
    L.rep(s, "Число пациентов, которых нужно лечить для предотвращения одного дополнительного события, также "
             "относится к конкретному исходу, сроку и группе сравнения.",
          "ЧБНЛ показывает, сколько пациентов нужно лечить за заданный срок, чтобы предотвратить одно "
          "событие по сравнению с плацебо. В FIDELIO-DKD для почечного исхода это 29 пациентов за 3 года. "
          "Проценты 17,8 и 21,1 относятся к медиане 2,6 года, поэтому число 29 из их разницы напрямую не "
          "получить.", notes=True)
    L.rep(s, "Улучшение промежуточного показателя, например альбуминурии, само по себе ещё не доказывает "
             "снижение смертности или других клинических исходов.",
          "Промежуточный показатель, например АКО, меняется раньше, а пользу подтверждают клинические "
          "исходы: их мы разберём на следующих слайдах. Теперь посмотрим программу исследований финеренона.",
          notes=True)
    notes_add(
        s,
        visit("Доктор, в FIDELIO-DKD риск почечного исхода был ниже на 18% по сравнению с плацебо: у 17,8% "
              "пациентов на финереноне против 21,1% на плацебо за 2,6 года."),
        "Если врач спросит о компонентах исхода: «Исход составной: почечная недостаточность, устойчивое "
        "снижение рСКФ на 40% и больше, смерть от почечных причин. Данные по каждому компоненту есть в "
        "публикации; запрос по ним передам в медицинский отдел.»")


# ------------------------------------------------------------------ slide 15
def s15(prs):
    s = L.S(prs, 15)
    L.set_title(s, "Остаточный риск сохраняется даже на ингибиторе АПФ или БРА с канаглифлозином: первичный "
                   "исход у 11,1% за 2,6 года")

    # one grey card on the full width: definition + study description
    info = lib.card(s, 36, 76, 888, 126, name="Карточка_исследование")
    move_before(info, L.shape_named(s, "Определение"))
    d = L.shape_named(s, "Определение")
    geom(d, 48, 84, 864, 22)
    write(d, [[("Остаточный риск:", 16, True, BLACK),
               (" риск осложнений, который сохраняется на фоне лечения.", 16, False, BLACK)]])
    st = L.shape_named(s, "Исследование")
    geom(st, 48, 112, 864, 86)
    L.rep(s, "выраженной альбуминурией.", nb("АКО выше 300 мг/г."), count=1)
    L.rep(s, "удвоение креатинина", "удвоение уровня креатинина", count=1)

    # chart 8 pt lower, with an asterisk on the title; the right card is aligned with it
    ch = [sh for sh in s.shapes if sh.has_chart][0]
    geom(ch, 36, 208, 543, 224)
    run = ch.chart.chart_title.text_frame.paragraphs[0].runs[0]
    run.text = "Доля пациентов с первичным исходом*"
    card = L.shape_named(s, "Карточка_результат")
    geom(card, 603, 208, 306, 224)
    orr = L.shape_named(s, "Отношение рисков")
    geom(orr, 618, 216, 276, 20)
    write(orr, [[("Отношение рисков 0,70 (95% ДИ 0,59–0,82)", 16, False, BLACK)]])
    ci = L.shape_named(s, "Доверительный интервал")
    geom(ci, 618, 236, 276, 40)
    write(ci, [[("Риск ниже на 30% по сравнению с плацебо (относительное снижение).", 16, False, BLACK)]])
    ob = L.shape_named(s, "Наблюдение")
    geom(ob, 618, 288, 276, 60)
    write(ob, [[("Наблюдение:", 16, True, BLACK), (" медиана 2,6 года. На графике доля пациентов с событием за "
                                                   "всё исследование.", 16, False, BLACK)]])
    bd = L.shape_named(s, "Граница доказательств")
    geom(bd, 618, 358, 276, 70)
    write(bd, [[("Остаточный риск:", 16, True, BLACK),
                (" на канаглифлозине и ингибиторе АПФ или БРА первичный исход возник у ", 16, False, BLACK),
                ("11,1%", 24, True, BLACK), (" пациентов.", 16, False, BLACK)]])

    # bridge to finerenone on a violet plate
    br = L.shape_named(s, "Состав исхода")
    geom(br, 36, 440, 888, 30)
    fill(br, PURPLE)
    no_line(br)
    round_rect(br, 8)
    write(br, [[("Остаточный риск остаётся даже на ингибиторе АПФ или БРА с ингибитором НГЛТ2. У финеренона "
                 "другая мишень: МКР.", 16, False, WHITE)]], anchor="m", ml=12, mr=8)

    footers(s, star="* Финеренон в исследовании CREDENCE не применяли; результаты разных исследований напрямую "
                    "не сравнивают.",
            abbr="АПФ: ангиотензинпревращающий фермент; БРА: блокатор рецепторов ангиотензина II; НГЛТ2: "
                 "натрий-глюкозный котранспортёр 2 типа; ДИ: доверительный интервал; АКО: отношение альбумина "
                 "к креатинину мочи; МКР: минералокортикоидный рецептор.",
            src="Perkovic V, et al. N Engl J Med. 2019;380:2295–2306.")

    # ---- notes
    L.rep(s, "Финеренон в CREDENCE не изучали. Результаты этой работы не доказывают эффект комбинации с "
             "финереноном. Проценты из разных исследований нельзя напрямую использовать для сравнения "
             "препаратов.",
          "Остаточный риск остаётся даже на ингибиторе АПФ или БРА с ингибитором НГЛТ2. Поэтому нужен ещё "
          "один механизм защиты. Он связан с рецептором, который мы сейчас рассмотрим.", notes=True)
    notes_add(
        s,
        visit("Доктор, в исследовании CREDENCE у пациентов с ХБП и диабетом на ингибиторе АПФ или БРА с "
              "канаглифлозином почечный или сердечно-сосудистый исход за 2,6 года возник у 11,1%. Остаточный "
              "риск есть, и у финеренона другая мишень: МКР."),
        ask("А что даёт финеренон на таком фоне?",
            "В FIDELIO-DKD финеренон добавляли к ингибитору АПФ или БРА в максимально переносимой дозе: "
            "почечный исход возник у 17,8% против 21,1% на плацебо (ОР 0,82; 95% ДИ 0,73–0,93; медиана 2,6 "
            "года). В FIDELITY ингибиторы НГЛТ2 принимали 6,7% участников, и результаты в этой подгруппе "
            "согласуются с общими; рандомизированного сравнения комбинации там не было. Сочетание с "
            "эмпаглифлозином изучали в CONFIDENCE (800 пациентов, 180 дней): АКО ниже от исходного на 52% "
            "при комбинации, на 32% на финереноне, на 29% на эмпаглифлозине. Прямого сравнения с "
            "канаглифлозином нет."),
        ask("Почему 11,1%, а у вас 17,8%?",
            "У CREDENCE другой первичный исход (с сердечно-сосудистой смертью), другая группа (АКО выше 300 "
            "мг/г) и другое лечение. Эти числа не сравнивают."))


# ------------------------------------------------------------------ slide 16
PALETTE_MAP = {"0E8DAC": TEAL, "0C91AD": TEAL, "C2752A": ORANGE, "C97925": ORANGE, "287C55": GREEN_D}


def remap(sh, mapping):
    for c in sh._element.iter(qn("a:srgbClr")):
        if c.get("val") in mapping:
            c.set("val", mapping[c.get("val")])


def s16(prs):
    s = L.S(prs, 16)
    L.set_title(s, "Альдостерон через МКР задерживает натрий и выводит калий")
    L.rep(s, "Натриевый канал ENaC", "Натриевый канал", count=1)
    L.rep(s, "Калиевый канал ROMK", "Калиевый канал", count=1)
    drop(s, "Энергозависимость натрий-калиевого насоса")

    # text labels: black for the nucleus / receptor group, palette colours for the ion flows, 18 pt Bold
    for n in ("Текст: Регуляция генов", "Текст: Внутриклеточный рецептор", "Текст: Комплекс переходит в ядро"):
        sh = L.shape_named(s, n)
        lib.set_text_color(sh, BLACK)
        remap(sh, {"7851AD": BLACK})          # leftovers in paragraph defaults
    for n, col in (("Текст: Натрий", TEAL), ("Текст: Натриевый канал", TEAL), ("Текст: В кровь", TEAL),
                   ("Текст: Калий", ORANGE), ("Текст: Калиевый канал", ORANGE), ("Текст: В мочу", ORANGE),
                   ("Текст: Альдостерон", GREEN_D)):
        lib.set_text_color(L.shape_named(s, n), col, size=18, bold=True)
    # the flow arrows, channels and pointers follow the labels
    for sh in lib.iter_shapes(s.shapes):
        remap(sh, PALETTE_MAP)
    # nucleus, receptor, dashed arrows: grey #7F7F7F
    solid_fill_of(L.shape_named(s, "Ядро"), CARD, line=GREY)
    for sh in lib.find_shapes(s, "Ядерная оболочка"):
        recolor(sh, GREY_L)
    for nm in ("Цепь ДНК", "Пары оснований ДНК"):
        for sh in lib.find_shapes(s, nm):
            recolor(sh, GREY)
    solid_fill_of(L.shape_named(s, "Минералокортикоидный рецептор"), GREY, line=GREY)
    for nm in ("Комплекс гормона и рецептора направляется в ядро", "Регуляция генов влияет на белки переноса",
               "Регуляция числа и активности натриевых каналов", "Регуляция натрий-калиевых насосов"):
        line_color(L.shape_named(s, nm), GREY)
    solid_fill_of(L.shape_named(s, "Один натрий-калиевый насос в базолатеральной мембране"), CHEV, line=GREY)
    # background zones follow the palette of the reference
    solid_fill_of(L.shape_named(s, "Просвет трубочки"), "C1E5F5")
    solid_fill_of(L.shape_named(s, "Интерстиций со стороны крови"), "FBE3D6")
    solid_fill_of(L.shape_named(s, "Главная клетка: наружная мембрана"), CARD, line=GREY_L)
    for sh in lib.find_shapes(s, "Внутренняя граница мембраны"):
        recolor(sh, CHEV)
    for sh in lib.find_shapes(s, "Митохондрия: контекст клетки"):
        solid_fill_of(sh, CHEV, line=GREY)
    for sh in lib.find_shapes(s, "Кристы митохондрии"):
        recolor(sh, GREY)
    for sh in lib.find_shapes(s, "Липидный бислой"):
        recolor(sh, GREY_L)
    for sh in lib.find_shapes(s, "Полярная головка липида"):
        solid_fill_of(sh, GREY_L)

    lab = L.shape_named(s, "Текст: В кровь")
    geom(lab, y=336)
    footers(s, abbr="МКР: минералокортикоидный рецептор.")

    # ---- notes
    notes_prepend(s, "Начнём с нормы: как минералокортикоидный рецептор работает в здоровой почке.")
    L.rep(s, "Эти процессы объясняют, почему действие на рецептор связано и с обменом калия.",
          "Поэтому блокада МКР влияет и на калий, и его контролируют по простому правилу; оно будет на слайде "
          "о назначении. А сначала посмотрим, что происходит, когда МКР активирован слишком сильно.",
          notes=True)
    notes_add(
        s,
        visit("Доктор, альдостерон включает рецептор МКР, и тот управляет обменом натрия и калия в почке. "
              "Финеренон блокирует этот рецептор, поэтому калий проверяют по простому правилу."),
        ask("МКР есть только в почке?",
            "МКР есть в почках, сердце, сосудах и иммунных клетках. Что происходит при его избыточной "
            "активации, показывает следующий слайд."),
        ask("Насколько часто растёт калий?",
            "В FIDELITY (13 026 пациентов, медиана 3,0 года) нежелательные явления по гиперкалиемии были у "
            "14,0% на финереноне против 6,9% на плацебо; окончательная отмена из-за гиперкалиемии 1,7% против "
            "0,6%; случаев со смертельным исходом не было. Калий проверяют до начала, через 4 недели и далее "
            "периодически."))


def run(prs):
    for f in (s09, s10, s11, s12, s13, s14, s15, s16):
        f(prs)
