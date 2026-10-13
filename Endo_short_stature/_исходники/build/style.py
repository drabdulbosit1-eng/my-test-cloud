"""Визуальная система «Кривая роста»: цвета, шрифты, геометрия.

Красный не используется: им автор ставит правки.
"""
from pptx.util import Inches, Pt

FONT_HEAD = 'Cambria'
FONT_BODY = 'Calibri'

C = {
    'ink': '1B2E3C',      # заголовки, тёмный фон
    'text': '26323A',     # основной текст
    'muted': '66737C',    # подписи, источники
    'axis': '8C979E',
    'centile': 'A9B4BA',
    'band': 'E7EDF0',     # полоса 3–97-го центиля
    'height': '1F6F8B',   # рост пациента, акцент
    'weight': 'C08A1E',   # масса пациента
    'rule': 'D3DADE',     # линии таблиц, шкала
    'paper': 'FFFFFF',
    'paper2': 'F3F6F7',   # подложка таблицы
    'onink': 'FFFFFF',
    'onink2': 'B9C7D0',
}

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MX = Inches(0.6)                 # поля слева и справа
CONTENT_W = SLIDE_W - 2 * MX

TITLE_Y = Inches(0.42)
TITLE_H = Inches(1.05)
BODY_Y = Inches(1.65)
BODY_BOTTOM = Inches(5.95)
EVID_Y = Inches(6.02)
SRC_Y = Inches(6.3)
RULER_Y = Inches(6.86)

SIZE_TITLE = Pt(28)
SIZE_BODY = Pt(20)
SIZE_TABLE = Pt(15)
SIZE_EVID = Pt(12)
SIZE_SRC = Pt(10)

PARTS = [
    (1, 4, 'Вступление'),
    (5, 7, 'Кривая роста'),
    (8, 15, 'Гипотиреоз'),
    (16, 23, 'Кушинг'),
    (24, 33, 'Пубертат'),
    (34, 41, 'ИФР-1 и ИФРСБ-3'),
    (42, 43, 'Итог'),
]
TOTAL = 43
