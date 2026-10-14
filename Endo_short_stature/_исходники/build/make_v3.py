"""Версия 3 по правке автора: текст/ → текст_v3/.

Что делает:
- убирает слайды 2–4 и 41–43 (их нет в версии автора);
- перенумеровывает слайды под новые вставки о механизмах (5, 7, 16, 31, 40);
- снимает ссылки «слайд N» с экрана, а в заметках переводит номера на новые;
- убирает строки, которые автор удалил у кривых (старые слайды 5 и 8);
- переносит описание кривых из figures.py в поле figure.

python3 make_v3.py <папка _исходники>
"""
import glob
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(__file__))
from figures import FIGS

src = sys.argv[1]
DROP = {2, 3, 4, 41, 42, 43}
NEW = {5, 7, 16, 31, 40}          # места новых слайдов о механизмах


def build_map():
    old = [n for n in range(1, 44) if n not in DROP]
    m, new_n = {}, 0
    for o in old:
        new_n += 1
        while new_n in NEW:
            new_n += 1
        m[o] = new_n
    return m


MAP = build_map()
MAP_THANKS = 43                   # последний слайд «Спасибо» добавляется отдельно

# Видимые правки: ссылки на слайды и строки, которые автор убрал
VISIBLE = {
    5: {'body': []},
    8: {'body': []},
    11: {'body_replace': [('Малую стартовую дозу со слайда 10 советуют ради безопасности, о росте данных нет',
                           'Малую стартовую дозу советуют ради безопасности, о росте данных нет')]},
    14: {'body_replace': [('Центральный гипотиреоз бывает частью дефицита нескольких гормонов гипофиза: '
                           'о надпочечниках — слайд 13',
                           'Центральный гипотиреоз бывает частью дефицита нескольких гормонов гипофиза')]},
    26: {'cell_replace': [('вариант нормы, слайд 24', 'вариант нормы'),
                          ('в том числе врождённый, слайды 27–28', 'в том числе врождённый'),
                          ('в том числе синдром Тернера, слайд 30', 'в том числе синдром Тернера')]},
    29: {'body_replace': [('Низкий рост с медленным темпом и малой массой: возможна функциональная форма (слайд 26)',
                           'Низкий рост с медленным темпом и малой массой: возможна функциональная форма')]},
    31: {'body_replace': [('Эстрогены с 11–12 лет, если ФСГ повышен в двух измерениях подряд; дозы — слайд 32',
                           'Эстрогены с 11–12 лет, если ФСГ повышен в двух измерениях подряд')]},
    39: {'cell_replace': [('низкий для возраста (слайд 35)', 'низкий для возраста')], 'drop_last_col': [2.8, 2.6, 4.2]},
    40: {'drop_last_col': [3.4, 3.4, 2.6]},
    1: {'body': ['Сиддиков Абдлубосит Адхамович', 'взрослый и детский эндокринолог'],
        'notes_replace': [('Кривые на слайдах — схемы, а не данные конкретных детей.',
                           'Кривые построены на нормах ВОЗ, а точки на них — примеры, не данные конкретных детей.')]},
}

THANKS = {'n': 43, 'part': 'Итог', 'layout': 'thanks', 'title': 'Спасибо за внимание!',
          'body': ['@drsiddikov', 'telegram'], 'table': None, 'source': None, 'claims': [], 'notes': ''}

REF = re.compile(r'(слайд\w*)(\s+)(\d+(?:\s*(?:–|,|и)\s*\d+)*)')


def renumber_notes(text, n_old):
    def fix(m):
        nums = m.group(3)

        def one(mm):
            o = int(mm.group(0))
            if o in DROP:
                raise SystemExit(f'слайд {n_old}: в заметках ссылка на удалённый слайд {o}')
            return str(MAP[o])
        return m.group(1) + m.group(2) + re.sub(r'\d+', one, nums)
    return REF.sub(fix, text)


def figure_for(n_old):
    f = FIGS.get(n_old)
    if not f:
        return None
    f = json.loads(json.dumps(f))
    if f['type'] == 'thumbs':
        for it in f['items']:
            it['note'] = re.sub(r';\s*слайд\s*\d+$', '', it['note'])
    return f


out = os.path.join(src, 'текст_v3')
os.makedirs(out, exist_ok=True)
for f in sorted(glob.glob(os.path.join(src, 'текст', '*.json'))):
    data = json.load(open(f, encoding='utf-8'))
    res = []
    for s in data:
        n = s['n']
        if n in DROP:
            continue
        s = dict(s)
        s['old_n'] = n
        s['n'] = MAP[n]
        s.pop('evidence_line', None)
        v = VISIBLE.get(n, {})
        if 'body' in v:
            s['body'] = v['body']
        for a, b in v.get('body_replace', []):
            assert a in s['body'], (n, a)
            s['body'] = [b if x == a else x for x in s['body']]
        for a, b in v.get('cell_replace', []):
            rows = s['table']['rows']
            hit = False
            for r in rows:
                for i, c in enumerate(r):
                    if c == a:
                        r[i], hit = b, True
            assert hit, (n, a)
        if v.get('drop_last_col'):
            t = s['table']
            t['header'] = t['header'][:-1]
            t['rows'] = [r[:-1] for r in t['rows']]
            s['col_w'] = v['drop_last_col']
        for a, b in v.get('notes_replace', []):
            assert a in s['notes'], (n, a)
            s['notes'] = s['notes'].replace(a, b)
        fig = figure_for(n)
        if fig:
            s['figure'] = fig
        if s.get('notes'):
            s['notes'] = renumber_notes(s['notes'], n)
        res.append(s)
    if os.path.basename(f).startswith('06'):
        res = [THANKS]
    if res:
        name = os.path.basename(f)
        json.dump(res, open(os.path.join(out, name), 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
        print(name, [(s.get('old_n'), s['n']) for s in res])

# Новые слайды о механизмах: механизмы.json; replace_old заменяет старый слайд
mpath = os.path.join(src, 'механизмы.json')
if os.path.exists(mpath):
    mech = json.load(open(mpath, encoding='utf-8'))
    files = {os.path.basename(f): json.load(open(f, encoding='utf-8')) for f in glob.glob(os.path.join(out, '*.json'))}
    for m in mech:
        placed = False
        for name, data in files.items():
            for i, s in enumerate(data):
                if m.get('replace_old') and s.get('old_n') == m['replace_old']:
                    keep = {k: s[k] for k in ('n', 'part', 'old_n')}
                    data[i] = {**m, **keep}
                    placed = True
        if not placed:
            target = m['file']
            files[target].append(m)
    for name, data in files.items():
        data.sort(key=lambda d: d['n'])
        json.dump(data, open(os.path.join(out, name), 'w', encoding='utf-8'), ensure_ascii=False, indent=2)
    print('механизмы:', [m['n'] if 'n' in m else m.get('replace_old') for m in mech])
