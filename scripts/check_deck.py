"""Check a pptx for problems that show only when PowerPoint opens it, and optionally fix the proofing language.

usage: python3 check_deck.py deck.pptx [--fix-lang out.pptx]

Reports
  size      slide size and ratio
  lang      runs whose proofing language does not match their text (Hangul -> ko-KR, otherwise en-US);
            a wrong or missing tag makes PowerPoint underline the whole run in red
  bold      bold set on a run whose own typeface (named, or a theme reference) is a Light face: PowerPoint thickens
            the strokes into a weight the font does not have. A bold run with no typeface of its own, in a deck whose
            theme body font is Light, is listed as "bold?" and not counted: it may inherit another font
  fonts     embedded fonts; a subset saved with another deck holds only that deck's characters and garbles new text
            on a machine without the font
  picture   one image (a logo) placed at different sizes or positions on the master, layouts and slides;
            it jumps when the audience moves between a cover and a content slide
  sizes     point sizes set on the slides themselves, to spot near-duplicate values for one role; a second line
            lists the sizes the layouts and masters carry (text with no size of its own inherits one of those)
Exit 1 when lang, bold, fonts or picture has a finding in the input deck (also when --fix-lang wrote a fixed copy).
"""
import os, re, sys, zipfile
if __name__ == "__main__" and (len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help")): sys.exit(__doc__)
from lxml import etree

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
HANGUL = re.compile(r"[가-힣ㄱ-ㆎ]")
PART = re.compile(r"ppt/(slides|slideLayouts|slideMasters|notesSlides)/[^/]+\.xml$")


def want_lang(text):
    return ("ko-KR", "en-US") if HANGUL.search(text) else ("en-US", "ko-KR")


def theme_fonts(z):
    names = [n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n)]
    t = etree.fromstring(z.read(sorted(names)[0]))
    get = lambda tag: (t.find(f".//{A}{tag}/{A}latin").get("typeface"), t.find(f".//{A}{tag}/{A}ea").get("typeface"))
    return {"+mj": get("majorFont"), "+mn": get("minorFont")}


def typeface(el, theme):
    for tag in ("ea", "latin"):
        e = el.find(A + tag)
        if e is not None and e.get("typeface"):
            tf = e.get("typeface")
            for k, v in theme.items():
                if tf.startswith(k):
                    tf = v[1] if tag == "ea" else v[0]
            if tf:
                return tf
    return None


def main():
    path = sys.argv[1]
    fix = None
    if "--fix-lang" in sys.argv:
        i = sys.argv.index("--fix-lang")
        if i + 1 >= len(sys.argv): sys.exit(__doc__)
        fix = sys.argv[i + 1]
        if os.path.abspath(fix) == os.path.abspath(path): sys.exit("check_deck: --fix-lang needs an output path different from the input")
    z = zipfile.ZipFile(path)
    bad = 0

    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    sz = pres.find("{http://schemas.openxmlformats.org/presentationml/2006/main}sldSz")
    cx, cy = int(sz.get("cx")), int(sz.get("cy"))
    print(f"size   {cx / 914400:.2f} x {cy / 914400:.2f} in, ratio {cx / cy:.3f} ({'16:9' if abs(cx / cy - 16 / 9) < .01 else '4:3' if abs(cx / cy - 4 / 3) < .01 else 'other'})")

    emb = pres.find("{http://schemas.openxmlformats.org/presentationml/2006/main}embeddedFontLst")
    if emb is not None:
        names = [f.find("{http://schemas.openxmlformats.org/presentationml/2006/main}font").get("typeface") for f in emb]
        print("fonts  embedded:", ", ".join(names), "- remove unless PowerPoint just saved them for this deck")
        bad += 1

    theme = theme_fonts(z)
    sizes, inherited, out = set(), set(), {}
    for n in sorted(z.namelist()):
        if not PART.match(n):
            continue
        t = etree.fromstring(z.read(n))
        changed = False
        for r in t.iter(A + "r", A + "fld"):
            text = "".join(x.text or "" for x in r.iter(A + "t"))
            rp = r.find(A + "rPr")
            if not text.strip():
                continue
            lang, alt = want_lang(text)
            have = rp.get("lang") if rp is not None else None
            if have != lang:
                print(f"lang   {n}: {have} on {text[:30]!r}")
                bad += 1
                if fix:
                    if rp is None:
                        rp = r.makeelement(A + "rPr", {}); r.insert(0, rp)
                    rp.set("lang", lang); rp.set("altLang", alt); changed = True
        for e in list(t.iter(A + "rPr")) + list(t.iter(A + "defRPr")):
            if e.get("sz"):
                (sizes if n.startswith("ppt/slides/") else inherited).add(int(e.get("sz")) / 100)
            if e.get("b") == "1":
                tf = typeface(e, theme)
                if tf and "light" in tf.lower():
                    print(f"bold   {n}: bold on {tf}")
                    bad += 1
                elif tf is None and any("light" in (f or "").lower() for f in theme["+mn"]):
                    print(f"bold?  {n}: bold with no typeface of its own; the theme body font is {theme['+mn'][0]}")
        if changed:
            out[n] = etree.tostring(t, xml_declaration=True, encoding="UTF-8", standalone=True)

    places = {}
    for n in sorted(z.namelist()):
        if not re.match(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml$", n):
            continue
        rels_name = n.rsplit("/", 1)[0] + "/_rels/" + n.rsplit("/", 1)[1] + ".rels"
        rels = {r.get("Id"): r.get("Target").rsplit("/", 1)[-1] for r in etree.fromstring(z.read(rels_name))} if rels_name in z.namelist() else {}
        for pic in etree.fromstring(z.read(n)).iter("{http://schemas.openxmlformats.org/presentationml/2006/main}pic"):
            blip = pic.find(".//" + A + "blip")
            xf = pic.find(".//" + A + "xfrm")
            if blip is None or xf is None or xf.find(A + "off") is None:
                continue
            img = rels.get(blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"))
            o, e = xf.find(A + "off"), xf.find(A + "ext")
            box = tuple(round(int(v) / 914400, 2) for v in (o.get("x"), o.get("y"), e.get("cx"), e.get("cy")))
            places.setdefault(img, {}).setdefault(box, []).append(n.split("/")[-1])
    for img, boxes in places.items():
        if len(boxes) > 1:
            print(f"picture {img} at {len(boxes)} places (x, y, w, h in):", "; ".join(f"{b} {', '.join(v)}" for b, v in boxes.items()))
            bad += 1

    print("sizes ", ", ".join(f"{s:g}" for s in sorted(sizes)) or "none set", "pt on the slides")
    print("       ", ", ".join(f"{s:g}" for s in sorted(inherited)) or "none", "pt in layouts, masters and notes")
    if fix:
        with zipfile.ZipFile(fix, "w", zipfile.ZIP_DEFLATED) as w:
            for item in z.infolist():
                w.writestr(item, out.get(item.filename, z.read(item.filename)))
        print("wrote", fix, f"({len(out)} parts retagged)")
    sys.exit(1 if bad else 0)


main()
