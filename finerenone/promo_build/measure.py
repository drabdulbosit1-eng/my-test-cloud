from PIL import ImageFont
_F = {}
def font(size, bold=False):
    k = (size, bold)
    if k not in _F:
        _F[k] = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSansNarrow-%s.ttf" % ("Bold" if bold else "Regular"), size * 10)
    return _F[k]
def width(text, size, bold=False):
    return font(size, bold).getlength(text) / 10 * 1.03
def n_lines(runs, w, size):
    """runs: list of (text, bold). Greedy word wrap; returns line count."""
    words = []
    for t, b in runs:
        for i, part in enumerate(t.split(" ")):
            if part == "" : continue
            words.append((part, b))
    lines, cur = 1, 0.0
    sp = width(" ", size)
    for wd, b in words:
        ww = width(wd, size, b)
        if cur == 0:
            cur = ww
        elif cur + sp + ww <= w:
            cur += sp + ww
        else:
            lines += 1; cur = ww
    return lines
def para_lines(paras, w, size):
    return sum(n_lines(p, w, size) for p in paras)
