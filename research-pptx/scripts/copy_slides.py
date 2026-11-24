# -*- coding: utf-8 -*-
"""Copy slides between decks (shapes, pictures) and keep the slide order sane.

  python -I scripts/copy_slides.py <src.pptx> <s1,s2,..> <dst.pptx> <after>
      -> copies the source slides (1-based) into dst right after slide <after>, same layout name; pictures are
         re-embedded. dst is saved in place.
  python -I scripts/copy_slides.py move <dst.pptx> <from> <to>
      -> moves slide <from> to position <to>.

"""
import copy
import io
import sys

from lxml import etree
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def layout_named(prs, name):
    for layout in prs.slide_layouts:
        if layout.name == name:
            return layout
    raise KeyError(name)


def copy_slide(src_slide, dst_prs):
    layout = layout_named(dst_prs, src_slide.slide_layout.name)
    new = dst_prs.slides.add_slide(layout)
    for sh in list(new.shapes):              # drop the layout's empty placeholders
        sh._element.getparent().remove(sh._element)
    tree = new.shapes._spTree
    for el in src_slide.shapes._spTree:
        if etree.QName(el).localname in ("nvGrpSpPr", "grpSpPr"):
            continue
        tree.append(copy.deepcopy(el))
    # re-embed pictures: every r:embed / r:link pointing at an image part
    remap = {}
    for rid, rel in src_slide.part.rels.items():
        if rel.reltype == RT.IMAGE:
            image_part = dst_prs.part.package.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
            new_rid = new.part.relate_to(image_part, RT.IMAGE)
            remap[rid] = new_rid
    for el in tree.iter():
        for attr in (R + "embed", R + "link"):
            v = el.get(attr)
            if v in remap:
                el.set(attr, remap[v])
    return new


def move(prs, frm, to):
    lst = prs.slides._sldIdLst
    ids = list(lst)
    el = ids[frm - 1]
    lst.remove(el)
    lst.insert(to - 1, el)


def main():
    if sys.argv[1] == "move":
        prs = Presentation(sys.argv[2])
        move(prs, int(sys.argv[3]), int(sys.argv[4]))
        prs.save(sys.argv[2])
        print("moved", sys.argv[3], "->", sys.argv[4])
        return
    src, nums, dst, after = sys.argv[1], [int(x) for x in sys.argv[2].split(",")], sys.argv[3], int(sys.argv[4])
    s, d = Presentation(src), Presentation(dst)
    for i, n in enumerate(nums):
        copy_slide(s.slides[n - 1], d)
        move(d, len(d.slides), after + 1 + i)
    d.save(dst)
    print(f"copied {nums} from {src} into {dst} after slide {after}; now {len(d.slides)} slides")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
