"""Promo rework, slides 25-32 of base47.pptx (old numbering), source: promo/final_25_32.md.
Build: cd W && PYTHONPATH=W:W/promo python3 apply14.py base47.pptx promo/t_25.pptx impl_25_32
One function per slide: s25(prs) ... s32(prs); run(prs) calls them in order.

Block order accepted by the customer: 24 -> 26 -> 27 -> 25 -> 31 -> 32 -> 33 -> 28 -> 29 -> 30 -> 34
(notes links are written for this order; no slide numbers inside slide text or notes).
Decisions of the customer that override final_25_32.md:
 - slide 25: no KDIGO row and no 'level A' until checked; title variant without KDIGO;
 - slide 27: ARTS stays with the patient group in the caption; gynecomastia (FIDELITY 0.1% / 0.2%) only in notes;
 - slide 30: notes end with a link to the slide about safety and special patient groups;
 - slide 32: notes get the 'what about eGFR after the start' block (moved from slide 8)."""
import copy
import io
import re

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

import measure as M
import r14lib as L
from lib import (add_box, add_text, bar, card, chevron, delete_shape, get_rpr, iter_shapes,
                 style_rpr)

BLACK = "000000"
PURPLE = "7030A0"   # finerenone and titles only
GREY = "7F7F7F"     # placebo
CYAN = "0F9ED5"     # comparator / SGLT2 inhibitor
GREEN = "4EA72E"    # combination
TEAL = "156082"     # neutral plate
BLUE = "2C71A4"     # neutral plate / 'keep the dose'
ORANGE = "E97132"   # stop / temporary withdrawal
YELLOW = "FFD009"   # caution (black text)
CARD = "F2F2F2"
CHEV = "D9D9D9"
DGREY = "666666"
BODY_BOTTOM = 538   # bottom edge of the footer text in the source deck
NB = " "


# ----------------------------------------------------------------------------- helpers
def nb(text):
    """Non-breaking spaces between a number and its unit, and inside 'рСКФ ≥25', '2 типа', '13 026'."""
    t = re.sub(r"(\d) (типа|мг|мл|мм|мкмоль|недел|месяц|года|лет|пациент|дн)", r"\1" + NB + r"\2", text)
    t = re.sub(r"(рСКФ|АКО|калии|калий) ([≥≤<>])", r"\1" + NB + r"\2", t)
    t = re.sub(r"(\d) (\d{3})(?!\d)", r"\1" + NB + r"\2", t)
    t = re.sub(r"(1,73) (м²)", r"\1" + NB + r"\2", t)
    t = re.sub(r"(мл/мин/1,73) ", r"\1" + NB, t)
    t = re.sub(r"(\d) (ммоль)", r"\1" + NB + r"\2", t)
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


def by_name(slide, name):
    return [sh for sh in iter_shapes(slide.shapes) if sh.name == name]


def one(slide, name):
    r = by_name(slide, name)
    if len(r) != 1:
        raise KeyError(f"{name!r}: {len(r)} shapes")
    return r[0]


def recolor_text(sh, color, bold=None, size=None):
    for r in sh._element.iter(qn("a:r")):
        style_rpr(get_rpr(r), color=color, bold=bold, size=size)


def recolor_fill(sh, color):
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(color)


def recolor_line(sh, color, width=None):
    sh.line.color.rgb = RGBColor.from_string(color)
    if width is not None:
        sh.line.width = Pt(width)


def remap_colors(slide, mapping):
    """Replace srgbClr values (fills, lines, gradient stops, text) on the whole slide."""
    n = 0
    for el in slide.shapes._spTree.iter(qn("a:srgbClr")):
        v = el.get("val")
        if v in mapping:
            el.set("val", mapping[v])
            n += 1
    return n


def rich(sh, paras, size=None, color=BLACK):
    """Rewrite a text shape: paras = list of paragraphs, each a list of (text, bold) runs.
    Keeps the first run format of the shape (font, size); `size` overrides it, `color` recolours."""
    L.set_rich_lines(sh, paras, color=color)
    if size is not None:
        for r in sh._element.iter(qn("a:r")):
            style_rpr(get_rpr(r), size=size)


def lines_of(paras, w, size):
    return sum(M.n_lines(p, w, size) for p in paras)


def txt(slide, x, y, w, h, paras, name, anchor="t", spacing=0.9, after=None):
    """Text box; paras = list of paragraphs, each a list of runs (text, size, bold, color)."""
    return add_text(slide, x, y, w, h, paras, anchor=anchor, name=name, line_spacing=spacing,
                    space_after=after)


def plate(slide, x, y, w, h, color, text, name, size=17, bold=False, align="l", text_color="FFFFFF"):
    sh = bar(slide, x, y, w, h, color, text, size=size, name=name, bold=bold, align=align)
    if text_color != "FFFFFF":
        recolor_text(sh, text_color)
    return sh


def grey_card(slide, x, y, w, h, name):
    return card(slide, x, y, w, h, name=name)


def footers(slide, abbr, sources, abbr_size=10):
    """Rewrite 'Источники' (8 pt, bottom edge fixed, grows upwards) and 'Сокращения' (10 pt, sits
    above the sources). Both black. Returns the y above which body content has to end."""
    src = L.shape_named(slide, "Источники")
    L.set_lines(src, sources)
    n = sum(M.n_lines([(p, False)], 850, 8) for p in sources)
    h = max(16.0, round(n * 9.8 + 2, 1))
    geom(src, y=BODY_BOTTOM - h, h=h)
    try:
        ab = L.shape_named(slide, "Сокращения")
    except KeyError:
        ab = add_text(slide, 36, 494, 850, 13, [[(abbr, abbr_size, False, BLACK)]], name="Сокращения",
                      line_spacing=1.0)
    L.set_lines(ab, [abbr])
    na = M.n_lines([(abbr, False)], 850, abbr_size)
    ha = round(na * 12.4, 1)
    top = min(494.0, BODY_BOTTOM - h - 3 - ha)
    geom(ab, y=top, h=ha)
    for sh in (src, ab):
        for r in sh._element.iter(qn("a:r")):
            style_rpr(get_rpr(r), color=BLACK)
    print(f"  footers: body must end above y={top:.0f}")
    return top


def notes(slide, *paras):
    """Speaker notes: paragraphs separated by an empty paragraph (as in the source deck)."""
    L.set_notes(slide, "\n\n".join(paras))


def notes_paras(slide):
    """Non-empty paragraphs of the notes as a list of strings."""
    return [p.text for p in slide.notes_slide.notes_text_frame.paragraphs if p.text.strip()]


def check_fit(label, paras, w, size, h, spacing=0.9):
    """Warn when text does not fit (Arial Narrow metrics); line height = size * 1.15 * spacing."""
    n = lines_of(paras, w, size)
    need = n * size * 1.15 * spacing
    flag = "OK " if need <= h + 0.5 else "!! "
    print(f"  fit {flag}{label}: {n} lines, {need:.0f} pt of {h:.0f} pt")
    return need


def write(sh, paras, size=None, color=BLACK, spacing=None, after=None):
    """Rewrite text of an existing shape. paras: list of paragraphs, each a list of (text, bold)."""
    rich(sh, [[(nb(t), b) for t, b in p] for p in paras], size=size, color=color)
    for i, p in enumerate(sh.text_frame.paragraphs):
        if spacing is not None:
            p.line_spacing = spacing
        if after is not None and i < len(sh.text_frame.paragraphs) - 1:
            p.space_after = Pt(after)


def flush(sh, anchor=None):
    """Zero insets (the shape was an autoshape with default 7.2 pt insets, text started right of x=36)."""
    from pptx.enum.text import MSO_ANCHOR
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if anchor == "t":
        tf.vertical_anchor = MSO_ANCHOR.TOP


def bulletize(sh, indent=12):
    """Native bullets ('•') with a hanging indent for every paragraph of a text shape."""
    for p in sh._element.iter(qn("a:p")):
        ppr = p.find(qn("a:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("a:pPr"))
            p.insert(0, ppr)
        ppr.set("marL", str(int(indent * 12700)))
        ppr.set("indent", str(-int(indent * 12700)))
        for k in ppr.findall(qn("a:buChar")) + ppr.findall(qn("a:buNone")):
            ppr.remove(k)
        etree.SubElement(ppr, qn("a:buChar")).set("char", "•")


# ----------------------------------------------------------------------------- slide 25
def s25(prs):
    s = L.S(prs, 25)
    # title without the KDIGO half until the KDIGO row is checked
    L.set_title(s, nb("ADA рекомендует финеренон при диабете 2 типа и ХБП с альбуминурией"))

    delete_shape(one(s, "Контекст"))
    delete_shape(one(s, "Местная_инструкция"))
    delete_shape(one(s, "Местная_инструкция_фон"))

    # two columns: plate (colour by meaning) + grey card
    y_pl, h_pl, y_card, h_card = 84, 34, 122, 226
    L.rep(s, "Рекомендация 11.8, уровень доказательств A", "ADA 2026, рекомендация 11.8")
    L.rep(s, "Рекомендация 11.9, уровень доказательств B", "ADA 2026, рекомендация 11.9, уровень B")
    for col, color in (("36", PURPLE), ("492", GREEN)):
        pl = one(s, f"Полоса_{col}")
        recolor_fill(pl, color)
        geom(pl, y=y_pl, h=h_pl)
        geom(one(s, f"Карточка_{col}"), y=y_card, h=h_card)

    main_h, y_main = 118, y_card + 12
    y_div = y_main + main_h + 4
    y_cond = y_div + 8
    h_cond = y_card + h_card - 8 - y_cond
    for col in ("36", "492"):
        geom(one(s, f"Текст_{col}"), y=y_main, h=main_h)
        geom(one(s, f"Разделитель_{col}"), y=y_div)
        geom(one(s, f"Условие_{col}"), y=y_cond, h=h_cond)
    t36 = ("При ХБП с альбуминурией ADA рекомендует нестероидный антагонист МКР с доказанной пользой "
           "(финеренон): он снижает риск прогрессирования ХБП и сердечно-сосудистых событий.")
    t492 = ("Одновременное начало приёма финеренона и ингибитора НГЛТ2 можно рассмотреть у взрослых с "
            "диабетом 2 типа при АКО ≥100 мг/г и рСКФ 30–90 мл/мин/1,73 м² на фоне ингибитора АПФ или БРА.")
    c36 = [("Условие: ", True), ("рСКФ ≥25 мл/мин/1,73 м²; после начала лечения контролируют калий.", False)]
    c492 = [("Основание: ", True), ("исследование CONFIDENCE. К 180-му дню АКО снизилось от исходного на 52% "
                                    "при сочетании, на 32% на финереноне и на 29% на эмпаглифлозине.", False)]
    write(one(s, "Текст_36"), [[(t36, False)]], size=18)
    write(one(s, "Текст_492"), [[(t492, False)]], size=18)
    write(one(s, "Условие_36"), [[c36[0], (c36[1][0], False)]], size=16)
    write(one(s, "Условие_492"), [[c492[0], (c492[1][0], False)]], size=16)
    check_fit("11.8 text", [[(t36, False)]], 396, 18, main_h)
    check_fit("11.9 text", [[(t492, False)]], 396, 18, main_h)
    check_fit("11.8 cond", [[(c36[0][0], True), (c36[1][0], False)]], 396, 16, h_cond)
    check_fit("11.9 cond", [[(c492[0][0], True), (c492[1][0], False)]], 396, 16, h_cond)

    # closing black plate (conclusion)
    plate(s, 36, y_card + h_card + 18, 888, 44, BLACK,
          "Финеренон входит в рекомендации ADA при ХБП и диабете 2 типа с альбуминурией",
          "Итоговая_плашка", size=18, bold=True, align="c")

    footers(s,
            "ХБП: хроническая болезнь почек; рСКФ: расчётная скорость клубочковой фильтрации; АКО: отношение "
            "альбумина к креатинину мочи; АПФ: ангиотензинпревращающий фермент; БРА: блокатор рецепторов "
            "ангиотензина II; НГЛТ2: натрий-глюкозный котранспортёр 2 типа; МКР: минералокортикоидный "
            "рецептор; ADA: Американская диабетическая ассоциация.",
            ["American Diabetes Association Professional Practice Committee. 11. Chronic Kidney Disease and "
             "Risk Management: Standards of Care in Diabetes-2026. Diabetes Care. 2026;49(Suppl 1):S246–S260. "
             "Agarwal R, et al. N Engl J Med. 2025;393:533–543."])

    # notes: link in, story, link out, visit phrase, questions
    L.rep(s, "В рекомендации 11.8 с уровнем доказательств A речь идёт о нестероидном антагонисте "
             "минералокортикоидного рецептора с доказанной эффективностью при ХБП с альбуминурией и рСКФ "
             "не менее 25 мл/мин/1,73 м²",
          "Рекомендация 11.8 касается нестероидного антагониста МКР с доказанной эффективностью при ХБП с "
          "альбуминурией и рСКФ не менее 25 мл/мин/1,73 м²", notes=True)
    L.rep(s, "Основанием послужило снижение альбуминурии в CONFIDENCE. Польза такой схемы по клиническим "
             "исходам этим исследованием не установлена.",
          "Основанием послужило снижение альбуминурии в CONFIDENCE: к 180-му дню АКО снизилось от исходного "
          "на 52% при сочетании, на 32% на финереноне и на 29% на эмпаглифлозине.", notes=True)
    story = notes_paras(s)
    # the closing paragraph about the label is replaced by the question blocks below
    story = [p for p in story if not p.startswith("Рекомендация профессионального общества")]
    if len(story) != 2:
        raise KeyError(f"s25 notes: expected 2 story paragraphs, got {len(story)}")
    notes(s,
          "Мы разобрали, чем финеренон отличается от спиронолактона. Рекомендации называют именно "
          "нестероидный антагонист МКР с доказанной пользой. Посмотрим, что в них написано. " + story[0],
          story[1],
          "Рекомендации определили место финеренона. Теперь практический вопрос: как начать лечение, на "
          "каких анализах и с какой дозы.",
          "Фраза для визита: «Финеренон входит в рекомендации ADA для пациентов с диабетом 2 типа и ХБП с "
          "альбуминурией. У каких ваших пациентов с диабетом 2 типа альбуминурия сохраняется на фоне "
          "ингибитора АПФ или БРА?»",
          "Если врач спросит: «Есть ли исходы при одновременном старте?» Ответ: CONFIDENCE оценивал АКО за "
          "180 дней (II фаза), клинические исходы в нём не изучали. Рекомендация 11.9 (уровень B) основана на "
          "снижении альбуминурии и допускает одновременный старт.",
          "Если врач спросит: «Написано ли это в вашей инструкции?» Ответ: Показание Финеренон-Орвилле: ХБП, "
          "связанная с диабетом 2 типа. Рекомендации описывают место препарата в схеме лечения, показание они "
          "не меняют. Одновременный старт с ингибитором НГЛТ2 в инструкции не описан: это рекомендация ADA, "
          "решение принимает врач. Условия начала и противопоказания смотрим по инструкции.",
          "Если врач спросит: «Какой уровень доказательств у рекомендации?» Ответ: У 11.9 уровень B. По 11.8 "
          "уровень уточню у медицинского отдела.",
          "Если врач спросит: «А европейские рекомендации?» Ответ: Рекомендации ESC и ERA 2026 по "
          "сердечно-сосудистым болезням и ХБП мы сверяем, класс и уровень назову после проверки.")


# ----------------------------------------------------------------------------- slide 26
def s26(prs):
    s = L.S(prs, 26)
    L.set_title(s, "Финеренон и ингибитор НГЛТ2 действуют на разные мишени: вместе они снижают альбуминурию "
                   "сильнее, чем по отдельности")
    delete_shape(one(s, "Местная_инструкция_25"))

    # the empty band y 76-122 disappears: branches go up
    y_pl, h_pl, y_card, h_card = 92, 32, 128, 70
    for col, color in (("36", GREEN), ("492", PURPLE)):
        pl = one(s, f"Ветка_{col}")
        recolor_fill(pl, color)
        geom(pl, y=y_pl, h=h_pl)
        geom(one(s, f"Ветка_текст_фон_{col}"), y=y_card, h=h_card)
        geom(one(s, f"Решение_{col}"), y=y_card + 4, h=h_card - 8)
    L.rep(s, "Ингибитор НГЛТ2 показан и переносится", "Пациент принимает ингибитор НГЛТ2")
    d36 = "Ингибитор НГЛТ2 продолжают, финеренон добавляют к текущей терапии по показаниям."
    d492 = "Финеренон назначают на фоне ингибитора АПФ или БРА после проверки калия, рСКФ и противопоказаний."
    write(one(s, "Решение_36"), [[(d36, False)]])
    write(one(s, "Решение_492"), [[(d492, False)]])
    check_fit("branch 1", [[(d36, False)]], 396, 17, h_card - 8)
    check_fit("branch 2", [[(d492, False)]], 396, 17, h_card - 8)

    # KDIGO scheme
    head = one(s, "Схема_заголовок")
    write(head, [[("Схема KDIGO: препараты добавляют к базовой терапии по показаниям", True)]])
    geom(head, y=214)
    ys = (246, 290, 334)
    for nm, y in zip(("Слой_284", "Слой_328", "Слой_372"), ys):
        geom(one(s, nm), y=y)
    L.rep(s, "Финеренон: при диабете 2\u00a0типа, АКО ≥30\u00a0мг/г, рСКФ\u00a0≥25\u00a0мл/мин/1,73\u00a0м² и нормальном калии",
          nb("Финеренон: диабет 2 типа, АКО ≥30 мг/г, рСКФ ≥25 мл/мин/1,73 м², нормальный калий"))
    recolor_fill(one(s, "Слой_328"), CYAN)
    arrow = one(s, "Стрелка_добавление")
    geom(arrow, x=96, y=246, h=126)
    recolor_fill(arrow, CHEV)
    cap = one(s, "Подпись_добавление")
    geom(cap, y=246, h=126)
    recolor_text(cap, BLACK)

    # grey card with the data
    y_c, h_c = 384, 84
    grey_card(s, 36, y_c, 888, h_c, "Карточка_данные")
    p1 = [("CONFIDENCE (одновременное начало, рСКФ 30–90, АКО ≥100 мг/г): ", True),
          ("к 180-му дню АКО снизилось от исходного на 52% при сочетании, на 32% на финереноне, на 29% на "
           "эмпаглифлозине. Группы плацебо не было.", False)]
    p2 = [("FIDELITY: ", True), ("результаты согласовались в подгруппе пациентов на ингибиторе НГЛТ2 "
                                 "(877 из 13 026, 6,7%).", False)]
    txt(s, 46, y_c + 6, 868, h_c - 12,
        [{"runs": [(nb(t), 16, b, BLACK) for t, b in p1], "space_after": 4},
         {"runs": [(nb(t), 16, b, BLACK) for t, b in p2]}], "Текст_данные", anchor="m")
    check_fit("data card", [p1, p2], 868, 16, h_c - 12)

    footers(s,
            "ХБП: хроническая болезнь почек; рСКФ: расчётная скорость клубочковой фильтрации; АКО: отношение "
            "альбумина к креатинину мочи; АПФ: ангиотензинпревращающий фермент; БРА: блокатор рецепторов "
            "ангиотензина II; НГЛТ2: натрий-глюкозный котранспортёр 2 типа; KDIGO: Kidney Disease: Improving "
            "Global Outcomes, международные рекомендации по болезням почек.",
            ["KDIGO 2024. Kidney Int. 2024;105(4S):S117–S314. KDIGO 2022 Clinical Practice Guideline for "
             "Diabetes Management in CKD. Kidney Int. 2022;102(5S):S1–S127. Agarwal R, et al. N Engl J Med. "
             "2025;393:533–543. Agarwal R, et al. Eur Heart J. 2022;43:474–484."])

    # notes
    L.rep(s, "Финеренон и ингибиторы НГЛТ2 действуют на разные мишени. Поэтому",
          "В CONFIDENCE мы увидели, что финеренон и эмпаглифлозин вместе снижают альбуминурию сильнее, чем "
          "каждый по отдельности. Теперь посмотрим, как это выглядит в схеме лечения. Препараты действуют на "
          "разные мишени. Поэтому", notes=True)
    L.rep(s, " Если один препарат не переносится, другой не становится его автоматической заменой: различаются "
             "механизмы, изученные группы пациентов и условия применения.", "", notes=True)
    L.rep(s, "Эти критерии рекомендаций нужно отличать от буквального текста местного показания. Возможность "
             "сочетания и последовательность лечения врач определяет по диагнозу, анализам, переносимости и "
             "инструкциям препаратов.",
          "Показание по инструкции Финеренон-Орвилле: ХБП, связанная с диабетом 2 типа. KDIGO описывает более "
          "узкую группу: нужны альбуминурия и ингибитор АПФ или БРА в максимальной дозе. Порог калия для "
          "старта по инструкции: не выше 4,8 ммоль/л, до 5,0 с дополнительным контролем в первые 4 недели. "
          "Сочетание и последовательность врач определяет по анализам и переносимости.", notes=True)
    story = [p for p in notes_paras(s) if not p.startswith("«Финеренон и ингибиторы НГЛТ2 действуют")]
    if len(story) != 2:
        raise KeyError(f"s26 notes: expected 2 story paragraphs, got {len(story)}")
    notes(s,
          story[0],
          story[1],
          "Мы разобрали, как финеренон сочетается с ингибитором НГЛТ2. Теперь вопрос, который врачи задают "
          "часто: чем финеренон отличается от спиронолактона, который они знают дольше.",
          "Фраза для визита: «Финеренон и ингибитор НГЛТ2 действуют на разные мишени, врачу не нужно выбирать "
          "между ними. В исследовании CONFIDENCE при одновременном начале обоих препаратов альбуминурия "
          "снизилась на 52% от исходной, на одном финереноне на 32%, на одном эмпаглифлозине на 29%».",
          "Если врач спросит: «Можно заменить ингибитор НГЛТ2 финереноном?» Ответ: Препараты не "
          "взаимозаменяемы, у них разные механизмы, исследования и показания. Если ингибитор НГЛТ2 не "
          "подходит, финеренон назначают на фоне ингибитора АПФ или БРА.",
          "Если врач спросит: «Польза сочетания по исходам доказана?» Ответ: CONFIDENCE (II фаза) оценивал "
          "АКО за 180 дней, исходов в нём нет. По исходам финеренон изучали в FIDELITY: ингибитор НГЛТ2 "
          "получали 877 из 13 026 пациентов (6,7%), результаты в этой подгруппе согласовались с общими. "
          "Рандомизации по ингибитору НГЛТ2 в FIDELITY не было.",
          "Если врач спросит: «Не растёт ли калий при сочетании?» Ответ: В CONFIDENCE калий выше 5,5 ммоль/л "
          "хотя бы раз был у 15,1% на сочетании, у 18,8% на финереноне и у 9,7% на эмпаглифлозине. Между "
          "сочетанием и одним финереноном значимой разницы нет (p=0,85). Контроль калия сохраняется.",
          "Если врач спросит: «А арГПП-1?» Ответ: арГПП-1 действуют на обмен веществ, финеренон блокирует "
          "МКР. В FIDELITY арГПП-1 получали 944 пациента (7,2%), результаты в этой подгруппе согласовались с "
          "общими. Сочетание врач выбирает по показаниям.")


# ----------------------------------------------------------------------------- slide 27
def cell(slide, x, y, w, h, fill, paras, name, anchor="m", margin=10):
    """Rounded cell with text inside. paras: list of paragraphs, each a list of runs (text, size, bold, color)."""
    sh = add_box(slide, x, y, w, h, fill=fill, radius=8, name=name)
    add_text(slide, x, y, w, h, [{"runs": [(nb(t), sz, b, c) for t, sz, b, c in p]} for p in paras],
             anchor=anchor, name=name, shape=sh, line_spacing=0.9)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = Pt(margin)
    tf.margin_top = tf.margin_bottom = Pt(4)
    return sh


def s27(prs):
    s = L.S(prs, 27)
    L.set_title(s, nb("Финеренон отличается от спиронолактона строением, избирательностью и доказательствами "
                      "при ХБП и диабете 2 типа"))
    # old content goes: texts, cards, Servier pictures (a kidney and heart only on the left read as a message)
    for nm in ("Picture 5", "Почка_финеренон", "Сердце_финеренон", "Text_7", "Text_9", "Text_10", "Text_11",
               "Карточка_финеренон", "Карточка_спиронолактон"):
        delete_shape(one(s, nm))

    xf, xs, wc, xl, wl = 194, 564, 360, 36, 150
    for nm, x in (("Полоса_финеренон", xf), ("Полоса_спиронолактон", xs)):
        pl = one(s, nm)
        geom(pl, x=x, y=84, w=wc, h=30)
    recolor_fill(one(s, "Полоса_финеренон"), PURPLE)
    recolor_fill(one(s, "Полоса_спиронолактон"), CYAN)

    SZ = 17
    pf = lambda txt_, b=False, c=BLACK, sz=SZ: (txt_, sz, b, c)
    rows = [
        ("Строение",
         [[pf("Нестероидный антагонист МКР. Метаболиты неактивны.")]],
         [[pf("Стероидный антагонист МКР. Есть активные метаболиты.")]]),
        ("Избирательность",
         [[pf("Избирательно блокирует МКР.")]],
         [[pf("Действует и на рецепторы половых гормонов, отсюда антиандрогенное действие.")]]),
        ("Исходы при ХБП и диабете 2 типа",
         [[pf("FIDELITY, 13 026 пациентов, медиана наблюдения 3,0 года, по сравнению с плацебо: риск "
              "комбинированного почечного исхода ниже на "), pf("23%", True, PURPLE),
           pf(", сердечно-сосудистого ниже на "), pf("14%", True, PURPLE)]],
         [[pf("Доказательства получены при тяжёлой СН (RALES). Крупных исследований исходов при ХБП и "
              "диабете 2 типа нет.")]]),
        ("Калий: прямое сравнение, ARTS",
         [[("5,3%", 18, True, PURPLE), (" \u00a0гиперкалиемия", 14, False, BLACK)]],
         [[("12,7%", 18, True, CYAN), (" \u00a0гиперкалиемия", 14, False, BLACK)]]),
    ]
    y = 120
    gap = 8
    for i, (label_, fin, spi) in enumerate(rows):
        def need(paras, w):
            n = 0
            for p in paras:
                n += M.n_lines([(t, b) for t, sz, b, c in p], w - 20, max(sz for t, sz, b, c in p))
            return n
        n = max(need(fin, wc), need(spi, wc), M.n_lines([(label_, True)], wl - 20, 16))
        h = max(46, round(n * SZ * 1.15 * 0.9 + 16))
        cell(s, xl, y, wl, h, CHEV, [[(label_, 16, True, BLACK)]], f"Название_строки_{i + 1}")
        cell(s, xf, y, wc, h, CARD, fin, f"Финеренон_строка_{i + 1}")
        cell(s, xs, y, wc, h, CARD, spi, f"Спиронолактон_строка_{i + 1}")
        print(f"  row {i + 1}: {n} lines, h={h}, y={y}")
        y += h + gap
    # caption of the ARTS row: group of patients and phase stay on the slide
    cap = txt(s, xf, y - gap + 4, 730, 20,
              [[("ARTS: ", 14, True, BLACK),
                ("II фаза, 392 пациента с СН со сниженной ФВ ЛЖ и умеренной ХБП, 4 недели, p=0,048", 14, False,
                 BLACK)]], "Подпись_ARTS")
    y_line = y - gap + 4 + 20 + 12
    txt(s, 36, y_line, 888, 28,
        [[("С финереноном спиронолактон и эплеренон не назначают. При любом антагонисте МКР контролируют "
           "калий.", 16, True, BLACK)]], "Нижняя_строка")
    print(f"  bottom line ends at y={y_line + 28}")

    footers(s,
            "ХБП: хроническая болезнь почек; СН: сердечная недостаточность; ФВ ЛЖ: фракция выброса левого "
            "желудочка; МКР: минералокортикоидный рецептор; RALES, ARTS, FIDELITY: названия исследований.",
            ["Финеренон-Орвилле. Инструкция по медицинскому применению. Agarwal R, et al. Eur Heart J. "
             "2022;43:474–484. Pitt B, et al. Eur Heart J. 2013;34:2453–2463. Pitt B, et al. N Engl J Med. "
             "1999;341:709–717. Aldactone (spironolactone). Prescribing information. Kerendia. EMA assessment "
             "report, 2022. Gerisch M, et al. Drug Metab Dispos. 2018;46:1546–1555."])

    notes(s,
          "Мы разобрали, как финеренон сочетается с ингибитором НГЛТ2. Теперь вопрос, который врачи задают "
          "часто: чем финеренон отличается от спиронолактона, который они знают дольше.",
          "Финеренон и спиронолактон блокируют один и тот же минералокортикоидный рецептор. Финеренон "
          "нестероидный, избирательно блокирует МКР, его метаболиты неактивны. Спиронолактон стероидный: он "
          "действует и на рецепторы половых гормонов, с чем связаны антиандрогенные эффекты, и у него есть "
          "активные метаболиты. В FIDELITY гинекомастия отмечена у 0,1% пациентов на финереноне и у 0,2% на "
          "плацебо.",
          "Главное различие в доказательствах. Для финеренона есть крупные исследования исходов у пациентов с "
          "ХБП и диабетом 2 типа. В FIDELITY у 13 026 пациентов по сравнению с плацебо риск почечного исхода "
          "был ниже на 23% (отношение рисков 0,77; 95% ДИ 0,67–0,88), риск сердечно-сосудистого исхода ниже "
          "на 14% (отношение рисков 0,86; 0,78–0,95); медиана наблюдения 3,0 года. Для спиронолактона и эплеренона таких "
          "исследований при ХБП и диабете 2 типа нет.",
          "Прямое сравнение есть по калию. В ARTS (II фаза, 392 пациента с сердечной недостаточностью со "
          "сниженной фракцией выброса и умеренной ХБП, 4 недели) гиперкалиемия возникла у 5,3% на финереноне "
          "и у 12,7% на спиронолактоне (p=0,048). Пациенты в ARTS другие, исследование II фазы, поэтому на "
          "визите называйте его полностью: ARTS, II фаза, сердечная недостаточность и ХБП.",
          "Рекомендация ADA называет именно нестероидный антагонист МКР с доказанной пользой. Посмотрим, что "
          "в ней записано.",
          "Фраза для визита: «Финеренон нестероидный и избирательный, и только у него есть крупные "
          "исследования исходов у пациентов с ХБП и диабетом 2 типа. По калию его сравнивали со "
          "спиронолактоном в ARTS (II фаза, сердечная недостаточность и ХБП): гиперкалиемия 5,3% против "
          "12,7%. С финереноном спиронолактон не назначают».",
          "Если врач спросит: «Что лучше?» Ответ: Прямого сравнения по почечным и сердечно-сосудистым исходам "
          "нет, и мы его не заявляем. У пациента с ХБП и диабетом 2 типа довод за финеренон в том, что именно "
          "для этой группы есть исследования FIDELIO-DKD и FIGARO-DKD и показание в инструкции. Прямое "
          "сравнение есть по калию (ARTS).",
          "Если врач спросит: «У спиронолактона разве нет доказательств?» Ответ: Есть, в другой группе. В "
          "RALES (1663 пациента с тяжёлой СН и ФВ ЛЖ не выше 35%) спиронолактон снизил риск смерти на 30% по "
          "сравнению с плацебо (относительный риск 0,70; 95% ДИ 0,60–0,82). Гинекомастия или боль в груди "
          "отмечены у 10% мужчин на спиронолактоне и у 1% на плацебо. При ХБП и диабете 2 типа исследований "
          "исходов у спиронолактона нет.",
          "Если врач спросит: «На чём основана избирательность?» Ответ: По инструкции финеренон селективный "
          "нестероидный антагонист МКР. В лабораторных работах показана высокая избирательность финеренона и "
          "недостаток её у спиронолактона (Bärfacker L, et al. ChemMedChem 2012). Клинические данные по "
          "половым гормонам: гинекомастия в FIDELITY 0,1% и 0,2% на плацебо (все участники). В RALES у мужчин "
          "на спиронолактоне 10% против 1% на плацебо. Исследования и группы разные, прямого сравнения нет.",
          "Если врач спросит: «Какие дозы были в ARTS?» Ответ: Финеренон 2,5, 5 или 10 мг один раз в сутки "
          "либо 5 мг два раза в сутки, спиронолактон 25 или 50 мг в открытой группе. Цифра 5,3% относится ко "
          "всем дозам финеренона вместе. Зарегистрированная схема 10 и 20 мг.",
          "Если врач спросит: «А эплеренон?» Ответ: Эплеренон тоже стероидный. Исследований исходов при ХБП и "
          "диабете 2 типа у него нет, с финереноном его не назначают (инструкция).")


# ----------------------------------------------------------------------------- slide 28
def s28(prs):
    s = L.S(prs, 28)
    L.set_title(s, "Повышение калия при финереноне объясняется механизмом действия и управляется контролем калия")

    # scheme: colours from the palette (cyan -> teal, brown/orange -> orange, gold and violet -> grey)
    n = remap_colors(s, {"088FAA": TEAL, "0E8DAC": TEAL, "C77422": ORANGE, "C2752A": ORANGE,
                         "C79849": DGREY, "7950A5": DGREY,
                         "EFF8FC": "C1E5F5", "FDF3F2": "FBE3D6",
                         "B5323C": PURPLE})
    print(f"  recoloured {n} colour values")
    # labels of the scheme: not smaller than 14 pt (they are 14-17 pt already)

    # right column: old blocks go, three 'plate + grey card' blocks come
    for nm in ("Риск выше", "Более высокий исходный калий",
               "Острое заболевание, обезвоживание или лекарства, повышающие калий либо концентрацию финеренона",
               "Калий выше нормы: гиперкалиемия. Нужен контроль калия и функции почек до начала и во время лечения."):
        delete_shape(one(s, nm))
    x0, w0 = 606, 318
    pad = 10
    y = 80
    # block 1: who should be checked
    plate(s, x0, y, w0, 26, TEAL, "Кому проверять калий особенно внимательно", "Полоса_риск", size=16)
    items = ["Более высокий исходный калий", "Сниженная функция почек", "Более высокая альбуминурия",
             "Женский пол", "Приём β-адреноблокаторов", "Более молодой возраст"]
    also = "Также: острое заболевание, обезвоживание, лекарства, повышающие калий или концентрацию финеренона."
    n_list = 1 + len(items)
    h_list = round(n_list * 14 * 1.15 * 0.9 + 3)
    n_also = M.n_lines([(also, False)], w0 - 2 * pad, 14)
    h_also = round(n_also * 14 * 1.15 * 0.9 + 3)
    h1 = 6 + h_list + 6 + h_also + 6
    y1 = y + 28
    grey_card(s, x0, y1, w0, h1, "Карточка_риск")
    txt(s, x0 + pad, y1 + 6, w0 - 2 * pad, h_list,
        [[("Анализ FIDELIO-DKD:", 14, True, BLACK)]] + [[(t, 14, False, BLACK)] for t in items],
        "Текст_факторы_риска")
    from lib import add_line
    add_line(s, x0 + pad, y1 + 6 + h_list + 3, x0 + w0 - pad, y1 + 6 + h_list + 3, color="BFBFBF", width=1,
             name="Разделитель_риск")
    txt(s, x0 + pad, y1 + 6 + h_list + 6, w0 - 2 * pad, h_also, [[(also, 14, False, BLACK)]], "Текст_также")
    print(f"  card 1: list {n_list} lines, also {n_also} lines, h={h1}, ends y={y1 + h1}")

    # block 2: FIDELITY numbers
    y2 = y1 + h1 + 6
    plate(s, x0, y2, w0, 26, TEAL, "FIDELITY, медиана наблюдения 3,0 года", "Полоса_FIDELITY", size=16)
    y2c = y2 + 28
    h2 = 72
    grey_card(s, x0, y2c, w0, h2, "Карточка_FIDELITY")
    txt(s, x0 + pad, y2c + 4, w0 - 2 * pad, 16,
        [[("Финеренон", 14, True, PURPLE), (" и ", 14, False, BLACK), ("плацебо", 14, True, GREY)]],
        "Легенда_FIDELITY")
    xl, wl_, xr, wr_ = x0 + 6, 160, x0 + 6 + 166, 146
    txt(s, xl, y2c + 19, wl_, 34,
        [[("14,0%", 28, True, PURPLE), ("\u00a0и\u00a0", 18, False, BLACK), ("6,9%", 28, True, GREY)]], "Число_гиперкалиемия")
    txt(s, xl, y2c + 52, wl_, 16, [[("гиперкалиемия", 14, False, BLACK)]], "Подпись_гиперкалиемия")
    txt(s, xr, y2c + 19, wr_, 34,
        [[("1,7%", 28, True, PURPLE), ("\u00a0и\u00a0", 18, False, BLACK), ("0,6%", 28, True, GREY)]], "Число_отмена")
    txt(s, xr, y2c + 52, wr_, 16, [[("окончательная отмена", 14, False, BLACK)]], "Подпись_отмена")

    # block 3: rule
    y3 = y2c + h2 + 6
    plate(s, x0, y3, w0, 26, TEAL, "Правило контроля", "Полоса_правило", size=16)
    y3c = y3 + 28
    rule = [("Калий и рСКФ: ", True), ("до начала, через 4 недели и далее периодически. ", False),
            ("Калий выше 5,5 ммоль/л: ", True), ("препарат временно отменяют.", False)]
    n_rule = M.n_lines(rule, w0 - 2 * pad, 15)
    h3 = round(n_rule * 15 * 1.15 * 0.9 + 14)
    grey_card(s, x0, y3c, w0, h3, "Карточка_правило")
    txt(s, x0 + pad, y3c + 6, w0 - 2 * pad, h3 - 10, [[(nb(t), 15, b, BLACK) for t, b in rule]], "Текст_правило")
    print(f"  card 3: {n_rule} lines, h={h3}, ends y={y3c + h3}")

    footers(s,
            "МКР: минералокортикоидный рецептор; рСКФ: расчётная скорость клубочковой фильтрации; "
            "гиперкалиемия в FIDELITY: нежелательные явления «гиперкалиемия» и «повышение калия в крови» по "
            "сообщению исследователя.",
            ["Финеренон-Орвилле. Инструкция по медицинскому применению. Kerendia. Summary of Product "
             "Characteristics (EMA). Agarwal R, et al. Eur Heart J. 2022;43:474–484. Agarwal R, et al. J Am Soc "
             "Nephrol. 2022;33:225–237. Agarwal R, et al. N Engl J Med. 2025;393:533–543. Agarwal R, et al. "
             "J Am Coll Cardiol. 2026;87:772–784. Sørensen MV, et al. JCI Insight. 2019;4:e126910. Czogalla J, "
             "et al. Pflügers Arch. 2016;468:849–858."])

    # notes
    L.rep(s, "Повышение калия связано с той же мишенью", "Мы разобрали, как начинать лечение, когда проверять "
          "анализы и как корректировать дозу. Теперь объясним, откуда берётся повышение калия и какой оно "
          "величины. Повышение калия связано с той же мишенью", notes=True)
    L.rep(s, " Наблюдение о меньшем риске у участников, получавших ингибиторы НГЛТ2, не доказывает защитного "
             "действия сочетания: по этой терапии пациентов случайным образом не распределяли. В CONFIDENCE "
             "добавление эмпаглифлозина статистически значимо не снизило риск гиперкалиемии. Поэтому калий и "
             "функцию почек проверяют до начала и во время лечения.", "", notes=True)
    L.rep(s, " Некоторые взаимодействия дополнительно меняют концентрацию финеренона. Их причина связана с его "
             "превращением в организме.", "", notes=True)
    story = notes_paras(s)
    if len(story) != 2:
        raise KeyError(f"s28 notes: expected 2 story paragraphs, got {len(story)}")
    notes(s,
          story[0],
          story[1],
          "Калий зависит и от концентрации самого финеренона. Её меняют лекарства, которые влияют на фермент "
          "CYP3A4. Об этом следующий слайд.",
          "Фраза для визита: «Повышение калия при финереноне связано с механизмом действия, поэтому калий "
          "проверяют по расписанию: до начала, через 4 недели и далее периодически. В FIDELITY из-за "
          "гиперкалиемии окончательно прекратили лечение 1,7% пациентов на финереноне и 0,6% на плацебо за "
          "время наблюдения около 3 лет».",
          "Если врач спросит: «Правда, что ингибиторы НГЛТ2 снижают риск гиперкалиемии?» Ответ: В анализе "
          "FIDELIO-DKD у получавших ингибиторы НГЛТ2 или диуретики риск был ниже. Наблюдение сделано без "
          "рандомизации: пациентов на такую терапию случайно не распределяли. В CONFIDENCE калий выше "
          "5,5 ммоль/л хотя бы раз был у 15,1% на сочетании, у 18,8% на финереноне и у 9,7% на эмпаглифлозине; "
          "между сочетанием и финереноном значимой разницы нет (p=0,85). Контроль калия сохраняется.",
          "Если врач спросит: «А если смотреть по строгому порогу?» Ответ: В FIDELIO-DKD калий выше "
          "5,5 ммоль/л хотя бы раз отмечен у 21,4% на финереноне и 9,2% на плацебо, выше 6,0 у 4,5% и 1,4%. "
          "По протоколу исследования препарат при этом временно отменяли и возобновляли по 10 мг при "
          "калии ≤5,0.",
          "Если врач спросит: «А тяжёлая гиперкалиемия бывает?» Ответ: В FIDELITY госпитализация из-за "
          "гиперкалиемии у 0,9% на финереноне и 0,2% на плацебо, случаев со смертельным исходом не было ни в "
          "одной группе. После 4-го месяца калий в среднем вырос на 0,21 ммоль/л на финереноне и на 0,02 на "
          "плацебо и дальше держался стабильно. Гипокалиемия реже на финереноне: 1,1% против 2,3%.")


# ----------------------------------------------------------------------------- slide 29
def s29(prs):
    s = L.S(prs, 29)
    # flat scheme instead of the raster picture: everything but the title, the footer and the logo is rebuilt
    for nm in ("Рисунок 22", "Финеренон", "Фермент", "Продукты", "Ингибиторы", "Концентрация выше", "Риск",
               "Индукторы", "Концентрация ниже", "Эффект", "Замедление", "Ускорение", "Сильные ингибиторы",
               "Умеренные ингибиторы", "Индукторы правило", "Контроль дозы", "Грейпфрут"):
        delete_shape(one(s, nm))
    L.set_title(s, "Лекарства, которые подавляют или ускоряют CYP3A4, меняют концентрацию финеренона")
    geom(L.title_shape(s), x=36, y=20, w=888, h=54)
    flush(L.title_shape(s), "t")
    flush(one(s, "Источники"))
    logo = one(s, "Логотип Орвилле")
    geom(logo, x=898.2, y=500, w=55.5, h=34)

    # subtitle
    sub = one(s, "Одинаковая доза")
    sub_txt = ("При одинаковой дозе. Около 90% финеренона превращается в неактивные метаболиты с участием "
               "CYP3A4, около 10% с участием CYP2C8.")
    write(sub, [[(sub_txt, False)]], size=16)
    geom(sub, x=36, y=76, w=888, h=22)
    sub.text_frame.word_wrap = True
    flush(sub, "t")
    check_fit("subtitle", [[(sub_txt, False)]], 888, 16, 22)

    # process ribbon: drug -> enzyme -> inactive metabolites
    plate(s, 36, 106, 200, 36, PURPLE, "Финеренон", "Узел_финеренон", size=16, align="c")
    chevron(s, 244, 110, w=56, h=28, name="Шеврон_1", color=CHEV)
    plate(s, 308, 106, 300, 36, TEAL, "CYP3A4: печень и стенка кишечника", "Узел_CYP3A4", size=16, align="c")
    chevron(s, 616, 110, w=56, h=28, name="Шеврон_2", color=CHEV)
    plate(s, 680, 106, 244, 36, DGREY, "Неактивные метаболиты", "Узел_метаболиты", size=16, align="c")
    cap = txt(s, 484, 146, 320, 20, [{"runs": [("ингибиторы замедляют фермент, индукторы ускоряют", 14, False, DGREY)],
                                      "align": "c"}], "Подпись_шевроны")

    # lanes: plate + effect card (left) and rule card (right)
    def lane(y, color, plate_text, effect, rules, key):
        plate(s, 36, y, 300, 34, color, plate_text, f"Полоса_{key}", size=16, bold=True)
        cell(s, 36, y + 38, 300, 86, CARD, [[(effect[0], 16, True, BLACK), (effect[1], 16, False, BLACK)]],
             f"Эффект_{key}", anchor="m")
        grey_card(s, 348, y, 576, 124, f"Правила_фон_{key}")
        paras = [[(nb(t), 17, b, BLACK) for t, b in r] for r in rules]
        box = txt(s, 360, y + 10, 552, 104, paras, f"Правила_{key}", anchor="m", after=8)
        bulletize(box)
        n = sum(M.n_lines([(t, b) for t, b in r], 552 - 12, 17) for r in rules)
        print(f"  lane {key}: {n} lines, need {n * 17 * 1.15 * 0.9 + 8 * (len(rules) - 1):.0f} of 104")

    lane(172, ORANGE, "Ингибиторы: метаболизм замедляется",
         ("Концентрация финеренона в крови выше. ", "Риск гиперкалиемии может повышаться."),
         [[("Сильные ингибиторы, например кларитромицин и ритонавир: ", False), ("противопоказаны", True),
           (".", False)],
          [("Умеренные и слабые ингибиторы: калий контролируют при начале приёма или изменении дозы; при "
            "необходимости врач корректирует дозу финеренона.", False)]], "ингибиторы")
    lane(304, BLUE, "Индукторы: метаболизм ускоряется",
         ("Концентрация финеренона в крови ниже. ", "Лечебный эффект может ослабевать."),
         [[("Сильные и умеренные индукторы, например рифампицин: совместное применение ", False),
           ("не рекомендуется", True), (".", False)],
          [("Грейпфрут и его сок исключают на время лечения.", False)]], "индукторы")

    # closing black plate
    plate(s, 36, 438, 888, 40, BLACK, "Перед назначением нужен полный список лекарств, добавок и безрецептурных "
          "средств.", "Итоговая_плашка", size=17, bold=True, align="c")

    footers(s, "CYP3A4, CYP2C8: изоферменты цитохрома P450.",
            ["Финеренон-Орвилле. Инструкция по медицинскому применению (метаболизм, сильные ингибиторы CYP3A4). "
             "Kerendia. Summary of Product Characteristics, раздел 4.5 (умеренные и слабые ингибиторы, "
             "индукторы, грейпфрут)."])

    # notes
    L.rep(s, "Финеренон превращается в неактивные метаболиты. Основной путь метаболизма финеренона обеспечивает "
             "фермент CYP3A4 в печени и стенке кишечника.",
          "Мы видели, что калий зависит от концентрации финеренона. Её меняют лекарства, которые влияют на "
          "фермент CYP3A4. Финеренон превращается в неактивные метаболиты; около 90% приходится на CYP3A4 в "
          "печени и стенке кишечника, около 10% на CYP2C8. Абсолютная биодоступность финеренона 43,5%: часть "
          "дозы метаболизируется при первом проходе через стенку кишечника и печень.", notes=True)
    L.rep(s, "К ним относятся кларитромицин, кетоконазол, итраконазол и ритонавир.",
          "К ним относятся кларитромицин, кетоконазол, итраконазол, ритонавир и нелфинавир.", notes=True)
    story = [p for p in notes_paras(s) if not p.startswith("«Перед назначением финеренона нужен")]
    if len(story) != 4:
        raise KeyError(f"s29 notes: expected 4 story paragraphs, got {len(story)}")
    notes(s,
          *story,
          "Теперь соберём все группы взаимодействий в одном слайде.",
          "Фраза для визита: «Перед назначением финеренона смотрим список лекарств: сильные ингибиторы "
          "CYP3A4, например кларитромицин, противопоказаны, а сильные и умеренные индукторы, например "
          "рифампицин, не рекомендуются».",
          "Если врач спросит: «Где это написано?» Ответ: Сильные ингибиторы указаны в инструкции "
          "Финеренон-Орвилле. Умеренные и слабые ингибиторы, индукторы и грейпфрут описаны в инструкции "
          "Kerendia (раздел 4.5). Конкретное сочетание и дозу оценивает врач.")


# ----------------------------------------------------------------------------- slide 30
def s30(prs):
    s = L.S(prs, 30)
    L.set_title(s, "Три группы взаимодействий финеренона: что противопоказано, чего избегают и что требует "
                   "контроля калия")
    geom(L.title_shape(s), y=20, h=54)

    # plates: colours from the palette (stop / caution / potassium control)
    recolor_fill(one(s, "Светофор_полоса_1"), ORANGE)
    p2 = one(s, "Светофор_полоса_2")
    recolor_fill(p2, YELLOW)
    recolor_text(p2, BLACK)
    recolor_fill(one(s, "Светофор_полоса_3"), TEAL)
    L.rep(s, "Не назначают совместно", "Избегают сочетаний")
    # icons: 'stop' stays white on orange; 'attention' = black triangle with a centred yellow '!'; 'K' in teal
    tri = one(s, "Значок_2")
    recolor_fill(tri, BLACK)
    tri.text_frame.text = ""
    gx, gy, gw, gh = (Emu(tri.left).pt, Emu(tri.top).pt, Emu(tri.width).pt, Emu(tri.height).pt)
    mark = add_text(s, gx, gy + 6.5, gw, 18, [{"runs": [("!", 16, True, YELLOW)], "align": "c"}], anchor="m",
                    name="Значок_2_знак")
    kk = one(s, "Значок_3")
    recolor_text(kk, TEAL, bold=True, size=16)

    # cards: shorter texts, 18 pt
    write(one(s, "Светофор_текст_1"),
          [[("Сильные ингибиторы CYP3A4: ", True),
            ("кларитромицин, кетоконазол, итраконазол, ритонавир, нелфинавир и др.", False)],
           [("Концентрация финеренона повышается.", False)]], size=18, after=8)
    write(one(s, "Светофор_текст_2"),
          [[("Другие антагонисты МКР: ", True), ("спиронолактон, эплеренон.", False)],
           [("Калийсберегающие диуретики: ", True), ("амилорид, триамтерен.", False)],
           [("Сильные и умеренные индукторы CYP3A4: ", True), ("рифампицин, карбамазепин, зверобой, эфавиренз.",
                                                              False)],
           [("Грейпфрут и его сок исключают.", False)]], size=18, after=8)
    write(one(s, "Светофор_текст_3"),
          [[("Препараты калия.", False)],
           [("Триметоприм, в том числе ко\u2011тримоксазол.", False)],
           [("Умеренные и слабые ингибиторы CYP3A4.", False)]], size=18, after=8)
    for k in (1, 2, 3):
        check_fit(f"card {k}", [[(sh_t, False)] for sh_t in
                                [one(s, f"Светофор_текст_{k}").text_frame.text]], 264, 18, 242)

    # bottom line becomes a grey card
    bot = one(s, "Светофор_вывод")
    write(bot, [[("Фоновая терапия: ", True),
                 ("к ингибитору АПФ или БРА в максимально переносимой дозе финеренон добавляют по показаниям. "
                  "Калий и функцию почек продолжают контролировать.", False)]], size=16)
    geom(bot, x=46, y=384, w=868, h=40)
    bot.text_frame.vertical_anchor = 3  # middle
    bg = grey_card(s, 36, 380, 888, 48, "Карточка_фоновая_терапия")
    bot._element.addprevious(bg._element)

    footers(s,
            "АПФ: ангиотензинпревращающий фермент; БРА: блокатор рецепторов ангиотензина II; МКР: "
            "минералокортикоидный рецептор; CYP3A4: изофермент цитохрома P450; KDIGO: Kidney Disease: Improving "
            "Global Outcomes, международные рекомендации по болезням почек.",
            ["Финеренон-Орвилле. Инструкция по медицинскому применению. Kerendia. Summary of Product "
             "Characteristics, раздел 4.5. KDIGO 2024. Kidney Int. 2024;105(4S):S117–S314."])

    # notes
    L.rep(s, "Взаимодействия различаются по требуемым действиям.",
          "Мы разобрали, как лекарства меняют концентрацию финеренона. Теперь соберём все группы "
          "взаимодействий в одном слайде. Взаимодействия различаются по требуемым действиям.", notes=True)
    L.rep(s, "По предоставленной местной инструкции финеренон противопоказано назначать вместе с сильными "
             "ингибиторами CYP3A4. К ним относятся кларитромицин, кетоконазол, итраконазол и ритонавир.",
          "Сильные ингибиторы CYP3A4 противопоказаны (инструкция Финеренон-Орвилле). К ним относятся "
          "кларитромицин, кетоконазол, итраконазол, ритонавир и нелфинавир.", notes=True)
    L.rep(s, "Европейская инструкция не рекомендует сильные и умеренные индукторы CYP3A4, предусматривает "
             "исключение грейпфрута и его сока и дополнительный контроль при умеренных или слабых ингибиторах. "
             "Калийсодержащие заменители соли также рассматриваются в европейском документе. Контроль при НПВП "
             "относится к клиническим рекомендациям. При их приёме врач оценивает и функцию почек.",
          "Инструкция Kerendia (ЕС) не рекомендует сильные и умеренные индукторы CYP3A4, предусматривает "
          "исключение грейпфрута и его сока и дополнительный контроль при умеренных или слабых ингибиторах.",
          notes=True)
    story = notes_paras(s)
    if len(story) != 3:
        raise KeyError(f"s30 notes: expected 3 story paragraphs, got {len(story)}")
    notes(s,
          *story,
          "Взаимодействия проверены. Осталось посмотреть, что нужно знать об особых группах пациентов: это "
          "следующий слайд, о безопасности.",
          "Фраза для визита: «Перед назначением финеренона врач проверяет список лекарств. Сильные "
          "ингибиторы CYP3A4 противопоказаны, другие антагонисты МКР и калийсберегающие диуретики не "
          "назначают, с препаратами калия и триметопримом контролируют калий».",
          "Если врач спросит: «Где это написано?» Ответ: Сильные ингибиторы, другие антагонисты МКР, "
          "калийсберегающие диуретики, препараты калия и триметоприм указаны в инструкции Финеренон-Орвилле. "
          "Индукторы, грейпфрут, умеренные и слабые ингибиторы описаны в инструкции Kerendia (раздел 4.5).",
          "Если врач спросит: «Что делать с НПВП и заменителями соли?» Ответ: В тексте инструкции "
          "Финеренон-Орвилле их нет. Общий клинический принцип для пациентов с ХБП: НПВП и калийсодержащие "
          "заменители соли могут повышать калий, а НПВП ещё и ухудшать функцию почек. Конкретную тактику "
          "выбирает врач, источник уточню у медицинского отдела.")


# ----------------------------------------------------------------------------- slide 31
def s31(prs):
    s = L.S(prs, 31)
    L.set_title(s, "Правило старта простое: калий определяет, можно ли начать, рСКФ определяет дозу 10 или 20 мг")

    # indication: plate 'Показание' + grey card, the text stays verbatim (without the word 'Показание:')
    pop = one(s, "Популяция")
    full = pop.text_frame.text
    if not full.startswith("Показание:"):
        raise KeyError("s31: indication text changed")
    verbatim = full[len("Показание:"):].strip()
    write(pop, [[(verbatim, False)]], size=15)
    geom(pop, x=48, y=104, w=864, h=64)
    pop.text_frame.vertical_anchor = 3  # middle
    bg = grey_card(s, 36, 104, 888, 64, "Карточка_показание")
    pop._element.addprevious(bg._element)
    plate(s, 36, 76, 120, 26, PURPLE, "Показание", "Полоса_показание", size=16)
    check_fit("indication", [[(verbatim, False)]], 864, 15, 64)

    # blocks go 10 pt down, rows are 6 pt lower, colours by meaning
    for nm in ("Калий_название", "Калий_единицы", "Доза_название", "Функция_почек_определение"):
        sh = one(s, nm)
        geom(sh, y=Emu(sh.top).pt + 10)
    rows = [(230, 48), (284, 72), (362, 48)]
    sets = [["Ячейка_калий_0", "Текст_калий_0", "Ячейка_рскф_0", "Текст_рскф_0", "Калий_до_4_8",
             "Начало_возможно", "Фильтрация_60", "Доза_20"],
            ["Ячейка_калий_1", "Текст_калий_1", "Ячейка_рскф_1", "Текст_рскф_1", "Калий_от_4_8_до_5",
             "Начало_с_контролем", "Фильтрация_25_60", "Доза_10"],
            ["Ячейка_калий_2", "Текст_калий_2", "Ячейка_рскф_2", "Текст_рскф_2", "Калий_больше_5",
             "Не_начинать_по_калию", "Фильтрация_ниже_25", "Не_начинать_по_фильтрации"]]
    for (y, h), names in zip(rows, sets):
        for nm in names:
            geom(one(s, nm), y=y, h=h)
    # potassium zones: customer palette (norm / keep with care / stop), doses: finerenone colour, stop
    recolor_fill(one(s, "Ячейка_калий_0"), "196B24")
    recolor_fill(one(s, "Ячейка_калий_1"), BLUE)
    recolor_fill(one(s, "Ячейка_калий_2"), ORANGE)
    recolor_fill(one(s, "Ячейка_рскф_0"), PURPLE)
    recolor_fill(one(s, "Ячейка_рскф_1"), PURPLE)
    recolor_fill(one(s, "Ячейка_рскф_2"), ORANGE)

    # text of the rows
    write(one(s, "Начало_возможно"), [[("Лечение можно начать.", False)]], size=17)
    L.rep(s, "Начало лечения можно рассмотреть при дополнительном контроле калия в первые 4 недели.",
          "Начало можно рассмотреть с учётом особенностей пациента при дополнительном контроле калия в первые "
          "4\u00a0недели.")
    for r in one(s, "Начало_с_контролем")._element.iter(qn("a:r")):
        style_rpr(get_rpr(r), size=15)
    check_fit("row 2 text", [[("Начало можно рассмотреть с учётом особенностей пациента при дополнительном "
                                "контроле калия в первые 4 недели.", False)]], 242, 15, 62)
    txt(s, 36, 414, 424, 16, [[("Условия старта: ", 14, True, BLACK),
                               (nb("рСКФ ≥25 и нет противопоказаний."), 14, False, BLACK)]], "Условия_старта")
    write(one(s, "Целевая_доза"), [[("Целевая и максимальная доза: ", True),
                                    ("20 мг один раз в сутки, независимо от приёма пищи.", False)]], size=18)

    footers(s, "рСКФ: расчётная скорость клубочковой фильтрации; CYP3A4: изофермент цитохрома P450.",
            ["Финеренон-Орвилле. Инструкция по медицинскому применению (перевод с узбекского): разделы "
             "«Показания», «Способ применения и дозы», «Противопоказания»."])

    # notes
    L.rep(s, "Полное показание в переводе предоставленного текста инструкции Узбекистана звучит так:",
          "Рекомендации определили место финеренона. Теперь практическое правило: как начать лечение. "
          "Показание по инструкции Финеренон-Орвилле (перевод с узбекского):", notes=True)
    story = [p for p in notes_paras(s) if not p.startswith("«Финеренон-Орвилле показан взрослым")]
    if len(story) != 3:
        raise KeyError(f"s31 notes: expected 3 story paragraphs, got {len(story)}")
    notes(s,
          *story,
          "Допустим, лечение началось. Теперь о том, когда проверять калий и рСКФ дальше.",
          "Фраза для визита: «Правило старта простое: два анализа, калий и рСКФ. Калий решает, можно ли "
          "начинать, рСКФ определяет дозу, 20 или 10 мг. Целевая доза 20 мг один раз в сутки».",
          "Если врач спросит: «Почему при рСКФ 25–60 доза 10 мг?» Ответ: Начальную дозу инструкция выбирает по "
          "рСКФ. При умеренном нарушении функции почек (клиренс креатинина от 30 до 60 мл/мин) AUC финеренона "
          "по инструкции выше на 34–36%. Цель 20 мг, повышение дозы идёт по калию (слайд «Коррекция дозы»).",
          "Если врач спросит: «Можно ли начинать при калии 4,9?» Ответ: По инструкции при калии выше 4,8 и не "
          "выше 5,0 ммоль/л начало можно рассмотреть, с дополнительным контролем калия в первые 4 недели. При "
          "калии выше 5,0 лечение не начинают.")


# ----------------------------------------------------------------------------- slide 32
def s32(prs):
    s = L.S(prs, 32)
    L.set_title(s, "Калий и рСКФ проверяют до первой дозы, через 4 недели и далее периодически")

    # scale: labels black, the third stage becomes 'Далее' (the interval of 4 months is not on the slide)
    L.rep(s, "Каждые 4\u00a0месяца¹", "Далее")
    L.rep(s, "Периодический контроль", "Далее периодически")
    L.rep(s, "KDIGO 2024 предлагает каждые 4\u00a0месяца¹; при повышенном риске чаще.",
          "Далее калий контролируют периодически, с учётом уровня калия и особенностей пациента.")
    for nm in ("Срок_180", "Срок_480", "Срок_780"):
        recolor_text(one(s, nm), BLACK)

    # same left inset (12 pt) in all plates; text in the three cards: top, 10 pt inset
    for nm in ("Этап_180", "Этап_480", "Этап_780", "Внеплановый_контроль_название"):
        one(s, nm).text_frame.margin_left = Pt(12)
    for k in ("180", "480", "780"):
        t = one(s, f"Этап_текст_{k}")
        geom(t, y=182 + 10, h=92 - 14)
        t.text_frame.vertical_anchor = 1  # top
        for para in t.text_frame.paragraphs:
            para.line_spacing = 0.9
    L.rep(s, " Срок определяет врач.", "")

    # row about combination with an SGLT2 inhibitor goes up; new row about eGFR
    comb = one(s, "Контроль_при_комбинации")
    write(comb, [[("С ингибитором НГЛТ2 калий контролируют так же:", True),
                  (" в CONFIDENCE сочетание значимо не снижало риск гиперкалиемии.", False)]], size=16)
    geom(comb, y=434)
    cell(s, 40, 462, 880, 26, CARD, [[("В первые 4 недели рСКФ может немного снизиться, затем она "
                                        "стабилизируется.", 16, False, BLACK)]], "Строка_рСКФ", margin=12)

    footers(s, "рСКФ: расчётная скорость клубочковой фильтрации; НГЛТ2: натрий-глюкозный котранспортёр 2 типа.",
            ["Финеренон-Орвилле. Инструкция по медицинскому применению. Kerendia. Summary of Product "
             "Characteristics. Agarwal R, et al. J Am Coll Cardiol. 2026;87:772–784."])

    # notes
    L.rep(s, "Наблюдение планируют с самого начала лечения. По предоставленной местной инструкции калий и рСКФ "
             "определяют до первой дозы, затем через четыре недели после начала, возобновления лечения или "
             "повышения дозы.",
          "Мы разобрали, как начать лечение. Теперь о том, когда проверять анализы. Наблюдение планируют с "
          "самого начала лечения. Калий и рСКФ определяют до первой дозы, затем через четыре недели после "
          "начала, возобновления лечения или повышения дозы (инструкция).", notes=True)
    L.rep(s, "Интервал каждые четыре месяца в стабильной ситуации приведён в KDIGO. Его нельзя считать "
             "универсальным сроком при любом изменении состояния. ", "", notes=True)
    L.rep(s, "местная инструкция предусматривает", "инструкция предусматривает", notes=True)
    story = [p for p in notes_paras(s) if not p.startswith("«Калий и рСКФ проверяют до первой дозы")]
    if len(story) != 2:
        raise KeyError(f"s32 notes: expected 2 story paragraphs, got {len(story)}")
    if story[1].count("Срок определяет врач.") != 1:
        raise KeyError("s32 notes: 'Срок определяет врач.' must stay once")
    notes(s,
          *story,
          "Что делать, если калий вырос: алгоритм коррекции дозы на следующем слайде.",
          "Фраза для визита: «Контроль простой: калий и рСКФ до первой дозы, через 4 недели и далее "
          "периодически. При остром заболевании, обезвоживании или новом препарате анализ делают раньше».",
          "Если врач спросит: «Как часто потом?» Ответ: Инструкция: периодически, с учётом калия и "
          "особенностей пациента. В протоколах FIDELIO-DKD и FIGARO-DKD калий определяли на плановых визитах "
          "примерно каждые 4 месяца, KDIGO 2022 предлагает такой же интервал. Конкретный срок определяет врач.",
          "Если врач спросит: «После начала упала рСКФ, что делать?» Ответ: По инструкции начало лечения может "
          "привести к небольшому снижению рСКФ в первые 4 недели, затем она стабилизируется. В FIDELITY "
          "снижение рСКФ как нежелательное явление отмечено у 5,3% на финереноне и 4,2% на плацебо. Оценку и "
          "тактику выбирает врач.",
          "Если врач спросит: «Что с рСКФ после старта?» Ответ: В первые 4 недели рСКФ может немного "
          "снизиться, затем стабилизируется; калий и рСКФ проверяют через 4 недели. Решение о коррекции "
          "принимает врач.",
          "Если врач спросит: «С ингибитором НГЛТ2 калий можно контролировать реже?» Ответ: Сочетание значимо "
          "не снижало риск гиперкалиемии. В CONFIDENCE калий выше 5,5 ммоль/л хотя бы раз был у 15,1% на "
          "сочетании и у 18,8% на финереноне (p=0,85). Режим контроля тот же.")


# ----------------------------------------------------------------------------- run
def run(prs):
    s25(prs)
    s26(prs)
    s27(prs)
    s28(prs)
    s29(prs)
    s30(prs)
    s31(prs)
    s32(prs)
