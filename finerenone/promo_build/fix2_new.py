"""Checker fixes for the new slides H1-H6 (48-53 in the pre-reorder build).
Applied after impl_new / impl_01_08 ... impl_41_47 and before impl_final (which reorders and edits junction
notes). Slides are addressed by their pre-reorder numbers: H1 = 48 ... H6 = 53.

Build: cd W && PYTHONPATH=W:W/promo python3 apply14.py base47.pptx promo/c_new.pptx impl_new impl_01_08 \
       impl_09_16 impl_17_24 impl_25_32 impl_33_40 impl_41_47 fix2_new impl_final
"""
from pptx.util import Pt

import r14lib as L

NBH = "‑"  # non-breaking hyphen: keeps study names such as FIDELIO-DKD on one line


def _shape(slide, name):
    return L.shape_named(slide, name)


def h1(prs):
    s = L.S(prs, 48)
    # every percentage of an outcome carries the word "combined" (rule of final_new, section "Рецензия" 3)
    L.rep(s, "перенесли почечный исход", "перенесли комбинированный почечный исход", count=1)
    # the caption now needs 6 lines (126 pt): give it room and move the CREDENCE text down
    cap = _shape(s, "Подпись_плацебо_FIDELIO-DKD")
    cap.height = Pt(128)
    _shape(s, "Текст_CREDENCE").top = Pt(314)
    # footnote about albuminuria as an inclusion criterion read as a part of footnote 2 (pre-clinical data):
    # give it its own mark on the closing line about the patient group
    L.rep(s, "типа. Калий", "типа³. Калий", count=1)
    L.rep(s, "² Доклинические данные. В FIDELIO-DKD", "² Доклинические данные. ³ В FIDELIO-DKD", count=1)


def h3(prs):
    s = L.S(prs, 50)
    # footnote numbers follow the order of appearance on the slide: plate 2 first, caption below the chain second
    L.rep(s, "слабее²", "слабее¹", count=1)
    L.rep(s, "исходов³.", "исходов².", count=1)
    L.rep(s, "² Доклинические данные. ³ Апостериорный анализ объединённых данных FIDELIO-DKD и FIGARO-DKD.",
          "¹ Доклинические данные. ² Дополнительный анализ объединённых данных FIDELIO-DKD и FIGARO-DKD, "
          "не запланированный заранее.", count=1)
    # notes: the previous slide already ends with "the next slide shows the chain from the receptor to the
    # indication", so the opening keeps only the link back
    L.rep(s, "На предыдущем слайде финеренон блокировал МКР внутри клетки. Теперь соберём цепочку от этой блокады "
          "до строк показания.", "На предыдущем слайде финеренон блокировал МКР внутри клетки.", count=1, notes=True)
    # same wording in the notes: the term "post hoc" is replaced by a plain description
    L.rep(s, "В апостериорном анализе объединённых данных (12 512 пациентов) изменение",
          "В дополнительном анализе объединённых данных (12 512 пациентов), не запланированном заранее, изменение",
          count=1, notes=True)


def h4(prs):
    s = L.S(prs, 51)
    # footnote 2 is the only one on the slide: it becomes 1
    L.rep(s, "спиронолактоне²", "спиронолактоне¹", count=1)
    L.rep(s, "² ARTS, часть B", "¹ ARTS, часть B", count=1)
    # answer 1: reference of the percentages (change from baseline, as in the notes)
    L.rep(s, "АКО снизилось на", "АКО снизилось от исходного на", count=1)
    # answer 2: comparison group of the number needed to treat
    L.rep(s, "чтобы предотвратить 1 почечное", "чтобы по сравнению с плацебо предотвратить 1 почечное", count=1)
    # answer 3: study names must not be split at the hyphen
    L.rep(s, "изучены в FIDELIO-DKD и FIGARO-DKD", f"изучены в FIDELIO{NBH}DKD и FIGARO{NBH}DKD", count=1)
    # notes: no echo of the end of the previous slide ("Следующие слайды собирают возражения врачей"), NNT group
    L.rep(s, "Теперь вопросы, которые врач задаёт на визите. Первые четыре касаются других препаратов.",
          "Первые четыре возражения касаются других препаратов.", count=1, notes=True)
    L.rep(s, "чтобы предотвратить 1 почечное событие за 3 года, нужно лечить",
          "чтобы по сравнению с плацебо предотвратить 1 почечное событие за 3 года, нужно лечить",
          count=1, notes=True)


def h5(prs):
    s = L.S(prs, 52)
    # FIGARO-DKD also enrolled patients with ACR 300-5000 mg/g at eGFR >= 60: the slide must not read as "only these"
    L.rep(s, "FIGARO-DKD включал пациентов с умеренно", "FIGARO-DKD включал и пациентов с умеренно", count=1)
    L.rep(s, "включал пациентов 2–4-й стадии с умеренно повышенной альбуминурией и 1–2-й стадии с выраженной.",
          "включал и пациентов 2–4-й стадии с умеренно повышенной альбуминурией (АКО 30–300 мг/г), и пациентов "
          "1–2-й стадии с АКО 300–5000 мг/г при рСКФ 60 и выше.", count=1, notes=True)
    # answer 4: comparison group of the numbers needed to treat; study names on one line
    L.rep(s, "чтобы предотвратить 1 событие", "чтобы по сравнению с плацебо предотвратить 1 событие", count=1)
    L.rep(s, "почечный исход, FIDELIO-DKD)", f"почечный исход, FIDELIO{NBH}DKD)", count=1)
    L.rep(s, "сердечно-сосудистый исход, FIGARO-DKD)", f"сердечно-сосудистый исход, FIGARO{NBH}DKD)", count=1)
    L.rep(s, "Для оценки пользы: чтобы предотвратить 1 событие",
          "Для оценки пользы: чтобы по сравнению с плацебо предотвратить 1 событие", count=1, notes=True)
    # the abbreviation is no longer defined on the slide, so the notes spell it out
    L.rep(s, "Экономической оценки ЧБНЛ не содержит.",
          "Число пациентов, которых нужно лечить, не является экономической оценкой.", count=1, notes=True)


def run(prs):
    h1(prs)
    h3(prs)
    h4(prs)
    h5(prs)
