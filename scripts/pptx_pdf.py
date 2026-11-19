"""Render a .pptx deck with PowerPoint itself: PDF beside the deck, one PNG per slide, optional contact sheets, bounds check.

usage: python3 pptx_pdf.py deck.pptx [--dpi 80] [--grid] [--no-png]
  deck.pdf            PowerPoint's own PDF (what the reader sees)
  deck-pages/p-N.png  one image per slide (pdftoppm; N is zero-padded when the deck has 10 or more slides)
  deck-pages/grid-K.png  2x2 contact sheets with --grid (needs Pillow)
No fallback to another renderer: a failure is reported as a failure. macOS with Microsoft PowerPoint only.
"""
import argparse, glob, os, shutil, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))


def pdf(pptx_path):
    """powerpoint_pdf.sh: PowerPoint itself writes deck.pdf beside the deck."""
    src = os.path.abspath(pptx_path)
    r = subprocess.run(["bash", os.path.join(HERE, "powerpoint_pdf.sh"), src], capture_output=True, text=True, timeout=300)
    if r.stderr.strip(): print(r.stderr.strip())   # why it failed
    if r.returncode: raise SystemExit("PowerPoint export failed")
    return os.path.splitext(src)[0] + ".pdf"


def pages(pdf_path, dpi):
    out = os.path.splitext(pdf_path)[0] + "-pages"
    shutil.rmtree(out, ignore_errors=True); os.makedirs(out)
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf_path, os.path.join(out, "p")], check=True)
    return sorted(glob.glob(os.path.join(out, "p-*.png")), key=lambda p: int(p.rsplit("-", 1)[1].split(".")[0]))


def grid(pngs):
    from PIL import Image
    ims = [Image.open(p).convert("RGB") for p in pngs]; w, h = ims[0].size; out = []
    for k in range(0, len(ims), 4):
        g = Image.new("RGB", (w * 2, h * 2), "white")
        for i, im in enumerate(ims[k:k + 4]): g.paste(im, ((i % 2) * w, (i // 2) * h))
        p = os.path.join(os.path.dirname(pngs[0]), f"grid-{k // 4 + 1}.png"); g.save(p); out.append(p)
    return out


def bounds(pptx_path):
    """Shapes that leave the slide. Needs python-pptx; skipped without it."""
    try:
        from pptx import Presentation
    except ImportError:
        return ["(python-pptx missing: bounds not checked)"]
    p = Presentation(pptx_path); W, H = p.slide_width, p.slide_height; bad = []
    for i, s in enumerate(p.slides, 1):
        for sh in s.shapes:
            if sh.left is None: continue
            if sh.left < 0 or sh.top < 0 or sh.left + sh.width > W + 1 or sh.top + sh.height > H + 1:
                bad.append(f"slide {i}: '{sh.name}' leaves the slide")
    return bad


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("pptx"); ap.add_argument("--dpi", type=int, default=80)
    ap.add_argument("--grid", action="store_true"); ap.add_argument("--no-png", action="store_true")
    a = ap.parse_args()
    for line in bounds(a.pptx): print("bounds:", line)
    d = pdf(a.pptx); print(d)
    if not a.no_png:
        pngs = pages(d, a.dpi); print(f"{len(pngs)} pages in {os.path.dirname(pngs[0])}")
        if a.grid:
            try: print("\n".join(grid(pngs)))
            except ImportError: print("(Pillow missing: no contact sheets)")
