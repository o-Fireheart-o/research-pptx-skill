# -*- coding: utf-8 -*-
"""Read-only profile of a template / sample deck: what layouts exist, which slides use them, what is on a slide.

  python -I deck_profile.py <deck.pptx>              slide size, layouts (placeholders, use count), slide list
  python -I deck_profile.py <deck.pptx> 5,9-12       every shape on those slides (id, kind, name, box in inches,
                                                      fill / line / font, first 60 characters of text)

Use the first form to map layouts to roles (cover, index, divider, body ...), then the second form on one
representative slide per role to copy its conventions (positions, fonts, colours) into the template profile.
"""
import sys

from pptx import Presentation
from pptx.util import Emu

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def inch(v):
    return "-" if v is None else f"{Emu(v).inches:.2f}"


def box(s):
    return f"x{inch(s.left)} y{inch(s.top)} w{inch(s.width)} h{inch(s.height)}"


def one_line(text, n=60):
    t = " / ".join(x.strip() for x in text.splitlines() if x.strip())
    return t if len(t) <= n else t[: n - 1] + "…"


def srgb(el, path):
    c = el.find(path)
    return None if c is None else c.get("val")


def style(s):
    """fill / line / first-run font, read from the shape XML (theme colours are shown as 'scheme')."""
    out = []
    sp = s._element.find(f"{P}spPr")
    if sp is not None:
        fill = srgb(sp, f"{A}solidFill/{A}srgbClr")
        if fill:
            out.append(f"fill={fill}")
        elif sp.find(f"{A}solidFill/{A}schemeClr") is not None:
            out.append("fill=scheme")
        line = srgb(sp, f"{A}ln/{A}solidFill/{A}srgbClr")
        if line:
            out.append(f"line={line}")
    rpr = s._element.find(f".//{A}r/{A}rPr")
    if rpr is not None:
        if rpr.get("sz"):
            out.append(f"{int(rpr.get('sz')) / 100:g}pt")
        if rpr.get("b") == "1":
            out.append("bold")
        col = srgb(rpr, f"{A}solidFill/{A}srgbClr")
        if col:
            out.append(f"color={col}")
        font = rpr.find(f"{A}latin")
        if font is not None:
            out.append(f"font={font.get('typeface')}")
    return " ".join(out)


def kind(s):
    if s.is_placeholder:
        return f"placeholder[{s.placeholder_format.idx}:{str(s.placeholder_format.type).split('.')[-1].split(' ')[0]}]"
    tag = s._element.tag.split("}")[1]
    if tag == "sp":
        c = s._element.find(f"{P}nvSpPr/{P}cNvSpPr")
        return "textbox" if c is not None and c.get("txBox") == "1" else "shape"
    return {"pic": "picture", "cxnSp": "connector", "grpSp": "group", "graphicFrame": "table/chart"}.get(tag, tag)


def parse_slides(arg, n):
    out = set()
    for part in arg.split(","):
        a, _, b = part.partition("-")
        out.update(range(int(a), int(b or a) + 1))
    return [i for i in sorted(out) if 1 <= i <= n]


def title_of(slide):
    t = slide.shapes.title
    if t is not None and t.text_frame.text.strip():
        return one_line(t.text_frame.text)
    for s in slide.shapes:
        if s.has_text_frame and s.text_frame.text.strip():
            return one_line(s.text_frame.text)
    return ""


def overview(prs):
    print(f"slide size: {inch(prs.slide_width)} x {inch(prs.slide_height)} in, {len(prs.slides)} slides\n")
    used = {}
    for slide in prs.slides:
        used[slide.slide_layout.name] = used.get(slide.slide_layout.name, 0) + 1
    print("== layouts (name | slides using it)")
    for layout in prs.slide_layouts:
        print(f"- {layout.name} | {used.get(layout.name, 0)}")
        for ph in layout.placeholders:
            print(f"    {kind(ph)} '{ph.name}' {box(ph)} {style(ph)}")
    print("\n== slides (no | layout | pictures | first text)")
    for n, slide in enumerate(prs.slides, 1):
        pics = sum(1 for s in slide.shapes if s._element.tag == f"{P}pic")
        print(f"{n:>3} | {slide.slide_layout.name} | {pics} | {title_of(slide)}")


def detail(prs, numbers):
    slides = list(prs.slides)
    for n in numbers:
        slide = slides[n - 1]
        print(f"== slide {n} | layout '{slide.slide_layout.name}'")
        for s in slide.shapes:
            text = one_line(s.text_frame.text) if s.has_text_frame else ""
            print(f"  #{s.shape_id} {kind(s)} '{s.name}' {box(s)} {style(s)}" + (f' "{text}"' if text else ""))
        print()


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    prs = Presentation(sys.argv[1])
    if len(sys.argv) > 2:
        detail(prs, parse_slides(sys.argv[2], len(prs.slides)))
    else:
        overview(prs)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
