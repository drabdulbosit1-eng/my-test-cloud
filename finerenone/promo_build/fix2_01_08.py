"""Check-pass fixes for original slides 1-8 (base47 numbers), applied after impl_01_08 and before the reorder.
Every text edit raises KeyError when its target is not found."""
import lib
import r14lib as L
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.util import Pt

NB = "\u00a0"
BLACK = "000000"
POINTER = "303030"
# pointer colours outside the palette on the nephron schemes (slides 3-5, 7) -> one neutral colour
STRAY_LINES = {"7E919E", "1F4E79", "783686", "202020", "586875"}


# ---------------------------------------------------------------- helpers
def recolor_pointers(s):
    """Thin pointer lines/arrows of a slide: stray colours -> #303030."""
    n = 0
    for sh in lib.iter_shapes(s.shapes):
        if sh.shape_type != 9:
            continue
        ln = sh._element.find(qn("p:spPr")).find(qn("a:ln"))
        if ln is None:
            continue
        for c in ln.iter(qn("a:srgbClr")):
            if c.get("val") in STRAY_LINES:
                c.set("val", POINTER)
                n += 1
    return n


def labels_black(s, *names):
    for nm in names:
        lib.set_text_color(L.shape_named(s, nm), BLACK)


def shift(sh, dx=0, dy=0):
    sh.left, sh.top = sh.left + Pt(dx), sh.top + Pt(dy)


# ---------------------------------------------------------------- slide 1: title
def s01(prs):
    """The conclusion line must not be wider than the indication: it lists the five outcomes of the
    indication (sustained eGFR decline, ESKD, CV death, non-fatal MI, hospitalisation for HF) instead of
    «почечных и сердечно-сосудистых осложнений» (the stroke is not in the indication)."""
    s = L.S(prs, 1)
    sub = L.shape_named(s, "Клиническая тема")
    L.set_lines(sub, ["Снижает риск устойчивого снижения функции почек, терминальной стадии болезни почек, "
                      "сердечно-сосудистой смерти, нефатального инфаркта миокарда и госпитализации по поводу "
                      "сердечной недостаточности у взрослых с хронической болезнью почек, связанной с "
                      "сахарным диабетом 2" + NB + "типа"])
    # the line is 6 lines long: 16 pt (the wrap of «сердечно-сосудистой» stays whole) and the stack on the
    # left panel moves up so that it ends level with the structural formula (y 98-453)
    lib.set_text_color(sub, BLACK, size=16)
    name = L.shape_named(s, "Text 3")
    name.top = Pt(100)
    sub.top, sub.left, sub.width, sub.height = Pt(169), Pt(66), Pt(378), Pt(118)
    goals = L.shape_named(s, "Практическая цель обучения")
    goals.top = Pt(307)
    mark = L.shape_named(s, "Маркировка")
    mark.top = Pt(417)
    L.rep(s, "Сегодня разберём финеренон: препарат, который снижает риск почечных и сердечно-сосудистых "
             "осложнений у взрослых с хронической болезнью почек, связанной с сахарным диабетом 2 типа.",
          "Сегодня разберём финеренон: препарат для взрослых с хронической болезнью почек, связанной с "
          "сахарным диабетом 2 типа. Он снижает риск устойчивого снижения функции почек, терминальной "
          "стадии болезни почек, сердечно-сосудистой смерти, нефатального инфаркта миокарда и "
          "госпитализации по поводу сердечной недостаточности.", notes=True)


# ---------------------------------------------------------------- slide 2
def s02(prs):
    s = L.S(prs, 2)
    # keep «давление в целевом диапазоне.» together: no single word on the second line
    L.rep(s, "и давление в" + NB + "целевом диапазоне.", "и давление" + NB + "в" + NB + "целевом" + NB + "диапазоне.")


# ---------------------------------------------------------------- slide 3
def s03(prs):
    s = L.S(prs, 3)
    labels_black(s, "Метка_приносящая", "Метка_выносящая", "Метка_собирательная")
    recolor_pointers(s)
    # the end of the previous slide's notes already says «вернёмся в конце курса» and «вспомним, как работает почка»
    L.rep(s, "К двум вопросам вернёмся в конце курса. Чтобы понять, как диабет повреждает почку и где "
             "действуют препараты, вспомним, как почка работает.", "Начнём с клубочка.", notes=True)
    # next slide says the same: «Теперь о том, как каналец сообщает клубочку...» is kept there only as a recap
    # (see s04), here the transition stays last


# ---------------------------------------------------------------- slide 4
def s04(prs):
    s = L.S(prs, 4)
    labels_black(s, "Метка_выносящая", "Метка_проксимальный")
    recolor_pointers(s)
    # the previous slide ends with «Сейчас о том, как клубочек сам регулирует фильтрацию»
    L.rep(s, "Мы увидели, как кровь фильтруется и как каналец возвращает нужное. Теперь о том, как каналец "
             "сообщает клубочку, сколько соли до него дошло. ",
          "Мы увидели, как кровь фильтруется и как каналец возвращает нужное. ", notes=True)


# ---------------------------------------------------------------- slide 5
def s05(prs):
    s = L.S(prs, 5)
    labels_black(s, "Метка_приносящая")
    recolor_pointers(s)
    # blue label «иНГЛТ2 блокируют этот транспорт»: it touched the cards of the chain (gap 2 pt)
    lab = L.shape_named(s, "Метка_иНГЛТ2")
    lab.left, lab.width = Pt(385), Pt(126)
    # the previous slide's notes end with «на следующем слайде при диабете она сбивается»
    L.rep(s, "На прошлом слайде плотное пятно держало фильтрацию в норме. При диабете эта регуляция "
             "сбивается. При диабете в первичной моче",
          "На прошлом слайде плотное пятно держало фильтрацию в норме. При диабете в первичной моче",
          notes=True)
    # a claim stronger than the data (CONFIDENCE: albuminuria, no outcomes of the combination)
    L.rep(s, "Поэтому врачу не нужно выбирать между этими препаратами, они дополняют друг друга.",
          "Поэтому эти препараты можно сочетать.", notes=True)


# ---------------------------------------------------------------- slide 6 (appendix)
def s06(prs):
    s = L.S(prs, 6)
    # the block of plates, panel captions, photos and cards is 7 pt higher and the cards are 4 pt taller:
    # the third card (3 lines of 16 pt) was full to the edge and the line under the cards (y 475) pressed on it
    keep = {"Заголовок", "Источники", "Сокращения", "Логотип Орвилле"}
    for sh in s.shapes:
        if sh.name in keep or sh.name.startswith("ШИК-окраска"):
            continue
        if Pt(95) <= sh.top < Pt(472):
            shift(sh, dy=-7)
            if sh.name.startswith("Карточка"):
                sh.height = Pt(70)


# ---------------------------------------------------------------- slide 7
def s07(prs):
    s = L.S(prs, 7)
    recolor_pointers(s)
    # the previous slide ends with «Второй путь: воспаление и фиброз в канальцах и ткани между ними»
    L.rep(s, "Клубочек мы рассмотрели. Теперь ткань вокруг него. При диабете повреждение",
          "При диабете повреждение", notes=True)
    # «меньше событий» needs its comparison group
    L.rep(s, "Клиническую пользу показали исходы: меньше почечных и сердечно-сосудистых событий в "
             "FIDELIO-DKD и FIGARO-DKD.",
          "Клиническую пользу показали исходы: по сравнению с плацебо меньше почечных и "
          "сердечно-сосудистых событий в FIDELIO-DKD и FIGARO-DKD.", notes=True)


# ---------------------------------------------------------------- slide 8
def s08(prs):
    s = L.S(prs, 8)
    # «недостаточность» (16 pt bold) went 4 pt beyond the left margin of 36 pt: label and heart 6 pt to the right
    shift(L.shape_named(s, "heart_source_label"), dx=6)
    shift(L.shape_named(s, "Сердце"), dx=6)
    # the previous slide ends with «Почки при диабете страдают не одни. Дальше о связи сердца и почек.»
    L.rep(s, "Теперь о связи двух органов: сердце и почки поражаются вместе.",
          "Сердце и почки поражаются вместе.", notes=True)


def run(prs):
    s01(prs)
    s02(prs)
    s03(prs)
    s04(prs)
    s05(prs)
    s06(prs)
    s07(prs)
    s08(prs)
