import re
import xml.etree.ElementTree as ET

DOC = "tmp/fixer-0813/unpacked/word/document.xml"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"

# 1-based line numbers for each w:drawing start tag, from raw text
raw = open(DOC, encoding="utf-8").read()
lines_with_drawing = [i + 1 for i, l in enumerate(raw.splitlines()) if "<w:drawing" in l]
print("lines containing <w:drawing:", lines_with_drawing)

tree = ET.parse(DOC)
root = tree.getroot()

def qname(e):
    return e.tag

drawings = root.iter(f"{W}drawing")
n_total = 0
n_empty = 0
for d in drawings:
    n_total += 1
    children = list(d)
    par = None
    # ElementTree has no parent pointers; find via walk
    def find_parent(el, target, parent=None):
        if el is target:
            return parent
        for c in el:
            r = find_parent(c, target, el)
            if r is not None:
                return r
        return None
    p = find_parent(root, d)
    gp = find_parent(root, p) if p is not None else None
    ggp = find_parent(root, gp) if gp is not None else None
    empty = len(children) == 0
    if empty:
        n_empty += 1
    print(
        f"drawing: children={[qname(c) for c in children]} empty={empty} "
        f"parent={qname(p) if p is not None else None} "
        f"grandparent={qname(gp) if gp is not None else None} "
        f"greatgrandparent={qname(ggp) if ggp is not None else None}"
    )
    if p is not None and p.tag == f"{MC}Choice":
        req = p.get("Requires")
        sibs = [qname(s) for s in gp] if gp is not None else []
        print(f"    Choice Requires={req!r} children-of-AlternateContent={sibs} "
              f"has-Fallback={any(s.tag == f'{MC}Fallback' for s in gp) if gp is not None else '?'}")

print(f"total w:drawing={n_total} empty={n_empty}")

# raw-text count of the exact collapsed shape
m = re.findall(r'<mc:Choice Requires="wps">\s*<w:drawing/>', raw)
print('raw regex <mc:Choice Requires="wps"><w:drawing/> matches:', len(m))
m2 = re.findall(r"<w:drawing/>", raw)
print("raw <w:drawing/> occurrences:", len(m2))
