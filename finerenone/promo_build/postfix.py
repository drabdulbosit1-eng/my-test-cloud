"""Package-level fixes: drop orphan notes masters, make master/layout ids unique."""
import os
import re
import shutil
import sys
import zipfile

SRC, DST = sys.argv[1], sys.argv[2]
WORK = "postfix_tmp"
shutil.rmtree(WORK, ignore_errors=True)
zipfile.ZipFile(SRC).extractall(WORK)


def read(p):
    with open(os.path.join(WORK, p), encoding="utf-8") as f:
        return f.read()


def write(p, s):
    with open(os.path.join(WORK, p), "w", encoding="utf-8") as f:
        f.write(s)


# 1. collect every relationship target in the package
targets = set()
for root, _, files in os.walk(WORK):
    for fn in files:
        if not fn.endswith(".rels"):
            continue
        rel_dir = os.path.relpath(root, WORK)            # e.g. ppt/slides/_rels
        owner_dir = os.path.dirname(rel_dir)             # ppt/slides
        x = open(os.path.join(root, fn), encoding="utf-8").read()
        for m in re.finditer(r'<Relationship ([^>]*)/>', x):
            a = m.group(1)
            if 'TargetMode="External"' in a:
                continue
            t = re.search(r'Target="([^"]+)"', a).group(1)
            if t.startswith("/"):
                full = t.lstrip("/")
            else:
                full = os.path.normpath(os.path.join(owner_dir, t))
            targets.add(full.replace("\\", "/"))

# 2. remove unreferenced notes masters (and their private theme + rels)
ct = read("[Content_Types].xml")
removed = []
for root, _, files in os.walk(WORK):
    for fn in files:
        p = os.path.relpath(os.path.join(root, fn), WORK).replace("\\", "/")
        if re.search(r"notesMasters/notesMaster\d+\.xml$", p) and p not in targets:
            d = os.path.dirname(p)
            rels = f"{d}/_rels/{fn}.rels"
            # theme used only by this notes master
            theme = None
            if os.path.exists(os.path.join(WORK, rels)):
                rx = read(rels)
                m = re.search(r'Target="([^"]*theme[^"]*)"', rx)
                if m:
                    theme = os.path.normpath(os.path.join(d, m.group(1))).replace("\\", "/")
                os.remove(os.path.join(WORK, rels))
            os.remove(os.path.join(WORK, p))
            removed.append(p)
            ct = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(p), "", ct)
            if theme and os.path.exists(os.path.join(WORK, theme)):
                os.remove(os.path.join(WORK, theme))
                trels = f"{os.path.dirname(theme)}/_rels/{os.path.basename(theme)}.rels"
                if os.path.exists(os.path.join(WORK, trels)):
                    os.remove(os.path.join(WORK, trels))
                ct = re.sub(r'<Override PartName="/%s"[^>]*/>' % re.escape(theme), "", ct)
                removed.append(theme)
write("[Content_Types].xml", ct)
print("removed:", removed)

# 3. unique ids: sldMasterId values and every sldLayoutId value must be unique (>= 2147483648)
pres = read("ppt/presentation.xml")
prels = read("ppt/_rels/presentation.xml.rels")
rid2t = {}
for m in re.finditer(r"<Relationship ([^>]*)/>", prels):
    a = m.group(1)
    rid2t[re.search(r'Id="([^"]+)"', a).group(1)] = re.search(r'Target="([^"]+)"', a).group(1)

next_id = 2147483648
master_entries = re.findall(r'<p:sldMasterId [^>]*/>', pres)
new_pres = pres
for entry in master_entries:
    rid = re.search(r'r:id="([^"]+)"', entry).group(1)
    new_entry = re.sub(r'id="\d+"', f'id="{next_id}"', entry, count=1)
    new_pres = new_pres.replace(entry, new_entry, 1)
    next_id += 1
    mpath = "ppt/" + rid2t[rid]
    mx = read(mpath)

    def bump(m):
        global next_id
        s = re.sub(r'id="\d+"', f'id="{next_id}"', m.group(0), count=1)
        next_id += 1
        return s

    mx = re.sub(r"<p:sldLayoutId [^>]*/>", bump, mx)
    write(mpath, mx)
write("ppt/presentation.xml", new_pres)
print("ids assigned up to", next_id - 1)

# 4. zip ([Content_Types].xml first)
if os.path.exists(DST):
    os.remove(DST)
with zipfile.ZipFile(DST, "w", zipfile.ZIP_DEFLATED) as z:
    z.write(os.path.join(WORK, "[Content_Types].xml"), "[Content_Types].xml")
    for root, _, files in os.walk(WORK):
        for fn in files:
            full = os.path.join(root, fn)
            arc = os.path.relpath(full, WORK).replace("\\", "/")
            if arc == "[Content_Types].xml":
                continue
            z.write(full, arc)
shutil.rmtree(WORK)
print("written", DST)
