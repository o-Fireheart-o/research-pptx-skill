# -*- coding: utf-8 -*-
"""Widen text boxes whose swapped text no longer fits on the same number of lines.

  python -I scripts/fit_boxes.py <before.pptx> <after.pptx>   (after.pptx is edited in place)
Compares each text box of the two decks (same layout, text swapped by para_swap.py). When the new text's
longest line is wider than the box, the box grows to fit (rough width estimate).
Blue (0000FF) notes keep their right edge; other boxes keep their left edge and stay on the slide.
"""
import sys

from pptx import Presentation
from pptx.util import Emu, Pt

sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.abspath(__file__)))
import para_swap as ps

SLIDE_W = 960.0


def cw(c):
    if c == " ":
        return 0.29
    if ord(c) < 0x80:                      # measured average widths, in em
        return 0.62 if c.isupper() or c.isdigit() else 0.56
    return 1.0


def widest(shape):
    mx, blue = 0.0, False
    for p in shape.text_frame.paragraphs:
        size = 12.0
        for r in p.runs:
            if r.font.size:
                size = r.font.size.pt
            rp = r._r.find(ps.A + "rPr")   # read the XML directly: r.font.color would insert an empty <a:solidFill/>
            if rp is not None and rp.find(ps.A + "solidFill/" + ps.A + "srgbClr[@val='0000FF']") is not None:
                blue = True
        for seg in p.text.split("\x0b"):
            mx = max(mx, sum(cw(c) for c in seg) * size)
    return mx, blue


def main(before, after):
    a, b = Presentation(before), Presentation(after)
    n_fix = 0
    for n, (sa, sb) in enumerate(zip(a.slides, b.slides), 1):
        xa = [s for s in sa.shapes if s.has_text_frame]   # top level only (group children use group coordinates)
        xb = [s for s in sb.shapes if s.has_text_frame]
        for old, new in zip(xa, xb):
            if old.text_frame.text == new.text_frame.text:
                continue
            tf = new.text_frame
            margin = Emu(tf.margin_left).pt + Emu(tf.margin_right).pt
            w = Emu(new.width).pt
            need, blue = widest(new)
            have = w - margin
            if need <= have:
                continue
            new_w = need + margin + 6
            if new_w > SLIDE_W - 8:
                print(f"  ! slide {n} {new.name}: text needs {new_w:.0f}pt, wider than the slide; shorten it")
                new_w = SLIDE_W - 8
            x = Emu(new.left).pt
            if blue:
                x = x + w - new_w
            x = max(4.0, min(x, SLIDE_W - new_w - 4))
            new.left, new.width = Pt(x), Pt(new_w)
            n_fix += 1
            print(f"  slide {n} {new.name}: {w:.0f} -> {new_w:.0f}pt  {tf.text.splitlines()[0][:40]}")
    b.save(after)
    print(n_fix, "boxes widened")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main(sys.argv[1], sys.argv[2])
