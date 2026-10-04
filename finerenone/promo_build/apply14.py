"""Usage: python3 apply14.py SRC.pptx DST.pptx module1 [module2 ...]
Loads SRC, runs module.run(prs) for each module in order, saves DST."""
import importlib
import sys

from pptx import Presentation

src, dst, *mods = sys.argv[1:]
prs = Presentation(src)
for m in mods:
    importlib.import_module(m).run(prs)
    print("applied", m)
prs.save(dst)
print("saved", dst)
