"""fix2 для слайдов 17-24 (нумерация base47). Применяется после impl_*, до перестановки и impl_final.
Правки по результатам проверки promo/check_17_24.md."""
import copy

from pptx.util import Emu, Pt

import lib as L0
import r14lib as L

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _dy(sh, dy):
    sh.top = Emu(int(sh.top + dy * 12700))


def _dh(sh, dh):
    sh.height = Emu(int(sh.height + dh * 12700))


KALIY_FULL = "через 4 недели после начала, возобновления или повышения дозы"
KALIY_OLD = "через 4 недели после начала или повышения дозы"


def s18(prs):
    s = L.S(prs, 18)
    # единая формула правила по калию (analysis_promo_47.md, раздел 6): с возобновлением
    L.rep(s, KALIY_OLD, KALIY_FULL, count=1)
    L.rep(s, KALIY_OLD, KALIY_FULL, notes=True, count=1)


def s20(prs):
    s = L.S(prs, 20)
    # блок «Калий» на слайде 20 стоял на 9 pt выше, чем на слайде 21: выравниваем по слайду 21
    d = 9.2
    _dh(L.shape_named(s, "Карточка_результат"), d)
    # карточка стала выше: немного воздуха между абзацами «Результата» (3 pt -> 6 pt), чтобы пустота не копилась внизу
    txt = L.shape_named(s, "Текст_результат")
    for p in txt.text_frame.paragraphs[:2]:
        p.space_after = Pt(6)
    _dh(txt, 6)
    for nm in ("Полоса_калий", "Карточка_калий", "Текст_калий"):
        _dy(L.shape_named(s, nm), d)
    L.rep(s, KALIY_OLD, KALIY_FULL, count=1)
    L.rep(s, KALIY_OLD, KALIY_FULL, notes=True, count=1)
    # вывод авторов анализа о калии подаём как вывод авторов
    L.rep(s, "при регулярном контроле калия последствия сводились к минимуму",
          "а по выводу авторов анализа регулярный контроль калия и меры при повышении калия "
          "свели влияние гиперкалиемии к минимуму", notes=True, count=1)


def s21(prs):
    s = L.S(prs, 21)
    # ХБП и СН были в списке сокращений, но в тексте стояли полностью
    L.rep(s, "Хроническая болезнь почек и сахарный диабет 2 типа:", "ХБП и диабет 2 типа:", count=1)
    L.rep(s, "СН: сердечная недостаточность; ", "", count=1)
    L.rep(s, KALIY_OLD, KALIY_FULL, count=1)   # начинается в обычном прогоне, после жирного «Контроль:»
    L.rep(s, KALIY_OLD, KALIY_FULL, notes=True, count=1)


def s22(prs):
    s = L.S(prs, 22)
    L.rep(s, "; СН: сердечная недостаточность", "", count=1)
    L.rep(s, KALIY_OLD, KALIY_FULL, count=1)   # начинается в обычном прогоне, после жирного «Контроль:»
    L.rep(s, KALIY_OLD, KALIY_FULL, notes=True, count=1)


def s23(prs):
    s = L.S(prs, 23)
    # заголовок диаграммы: перенос строки лежал внутри одного a:t (в impl условие не сработало)
    ch = L.shape_named(s, "Chart")
    for r in ch.chart._chartSpace.iter("{%s}r" % A):
        t = r.find("{%s}t" % A)
        if t is not None and "\n" in (t.text or "") and r.getparent().getparent().tag == "{%s}rich" % C:
            p0 = r.getparent()
            first, second = t.text.split("\n", 1)
            t.text = first
            p1 = copy.deepcopy(p0)
            p1.find("{%s}r" % A).find("{%s}t" % A).text = second
            p0.addnext(p1)
            break
    else:
        raise KeyError("chart title with line break not found")


def s24(prs):
    s = L.S(prs, 24)
    # цвет комбинации #4EA72E (в диаграммах стоял 4BAA22)
    n = 0
    for nm in ("АКО на 180 день", "Калий выше 5,5"):
        cs = L.shape_named(s, nm).chart._chartSpace
        for el in cs.iter("{%s}srgbClr" % A):
            if el.get("val") == "4BAA22":
                el.set("val", "4EA72E")
                n += 1
    assert n == 2, n
    L.rep(s, KALIY_OLD, KALIY_FULL, count=1)           # карточка «Первичная точка»
    # фраза для визита: не только цифры, но и значимость, иначе читается как «комбинация безопаснее»
    L.rep(s, "на одном финереноне у 18,8%».", "на одном финереноне у 18,8%; различие незначимо (p=0,85)».",
          notes=True, count=1)
    # переход: без вывода «врачу не нужно выбирать»
    L.rep(s, "Что это значит для выбора препарата: финеренон и иНГЛТ2 дополняют друг друга, врачу не нужно "
             "выбирать между ними. Об этом следующий слайд.",
          "Что это значит для лечения: финеренон и иНГЛТ2 действуют на разные мишени, и по альбуминурии "
          "вместе они дали больший эффект, чем по отдельности. Как это выглядит в схеме лечения, показывает "
          "следующий слайд.", notes=True, count=1)


def run(prs):
    s18(prs)
    s20(prs)
    s21(prs)
    s22(prs)
    s23(prs)
    s24(prs)
