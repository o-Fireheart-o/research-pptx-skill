# -*- coding: utf-8 -*-
"""Swap text paragraph by paragraph in a hand-edited deck (layout, pictures and shapes untouched).

Used to make a translated edition from a finished deck. Paragraph keys: runs in the
paragraph's main format are plain, runs in another format are wrapped in <f1>..</f1>, line breaks "\n".

  python -I scripts/para_swap.py extract <src.pptx> <dict.json> hangul|latin
      -> lists paragraphs that contain Hangul (hangul) or Latin-only text (latin); keeps existing values
  python -I scripts/para_swap.py build <src.pptx> <dict.json> <out.pptx>
      -> replaces every paragraph whose key has a non-empty value; slide 1 (cover) is skipped.
         A key "<slide>|<paragraph>" applies to that slide only and wins over the plain key
         (for short lines such as "    클릭" that mean different things on different slides).
"""
import copy
import json
import os
import re
import sys

from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
TAG = re.compile(r"<f(\d+)>(.*?)</f\1>", re.S)
HANGUL = re.compile("[가-힣]")


def walk(shapes):
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from walk(s.shapes)
        else:
            yield s


def text_frames(slide):
    for s in walk(slide.shapes):
        if s.has_text_frame:
            yield s.text_frame
        if getattr(s, "has_table", False) and s.has_table:
            for row in s.table.rows:
                for c in row.cells:
                    yield c.text_frame


def fmt_key(r):
    rp = r.find(A + "rPr")
    if rp is None:
        return ""
    rp = copy.deepcopy(rp)
    for k in ("lang", "altLang", "dirty", "err", "spc", "kern", "noProof"):
        rp.attrib.pop(k, None)
    return etree.tostring(rp).decode()


def tokens(p):
    out = []
    for el in p:
        tag = etree.QName(el).localname
        if tag in ("r", "fld"):
            t = el.findtext(A + "t") or ""
            k = fmt_key(el)
            if out and out[-1][0] == "r" and out[-1][1] == k:
                out[-1] = ("r", k, out[-1][2] + t, out[-1][3] + [el])
            else:
                out.append(("r", k, t, [el]))
        elif tag == "br":
            out.append(("br", None, "\n", [el]))
    return out


def marked(p):
    toks = tokens(p)
    weight = {}
    for kind, k, t, _ in toks:
        if kind == "r":
            weight[k] = weight.get(k, 0) + len(t.strip())
    if not weight:
        return None, None, toks
    base = max(weight, key=weight.get)
    idx, parts = {}, []
    for kind, k, t, _ in toks:
        if kind == "br":
            parts.append("\n")
        elif k == base or not t.strip():
            parts.append(t)
        else:
            i = idx.setdefault(k, len(idx) + 1)
            parts.append(f"<f{i}>{t}</f{i}>")
    return "".join(parts), (base, idx), toks


def plain(s):
    return TAG.sub(lambda m: m.group(2), s)


def needs(s, mode):
    t = plain(s).strip()
    if mode == "hangul":
        return bool(HANGUL.search(t))
    return bool(re.search("[A-Za-z]", t)) and not HANGUL.search(t)


def paragraphs(prs):
    for n, slide in enumerate(prs.slides, start=1):
        if n == 1:
            continue
        for tf in text_frames(slide):
            for p in tf.paragraphs:
                yield n, p._p


def rebuild(p, new, base, idx, toks, lang):
    rpr_of = {}
    for kind, k, t, el in toks:
        if kind == "r" and k not in rpr_of:
            rp = el[0].find(A + "rPr")
            rpr_of[k] = copy.deepcopy(rp) if rp is not None else None
    by_i = {i: k for k, i in idx.items()}
    for kind, k, t, els in toks:
        for el in els:
            p.remove(el)
    end = p.find(A + "endParaRPr")
    pos = list(p).index(end) if end is not None else len(p)

    def add(text, k):
        nonlocal pos
        for j, line in enumerate(text.split("\n")):
            if j:
                br = etree.Element(A + "br")
                if rpr_of.get(k) is not None:
                    br.append(copy.deepcopy(rpr_of[k]))
                p.insert(pos, br)
                pos += 1
            if line:
                r = etree.Element(A + "r")
                if rpr_of.get(k) is not None:
                    rp = copy.deepcopy(rpr_of[k])
                    rp.set("lang", lang)
                    r.append(rp)
                etree.SubElement(r, A + "t").text = line
                p.insert(pos, r)
                pos += 1

    last = 0
    for m in TAG.finditer(new):
        add(new[last:m.start()], base)
        add(m.group(2), by_i.get(int(m.group(1)), base))
        last = m.end()
    add(new[last:], base)


def extract(src, dict_path, mode):
    prs = Presentation(src)
    old = json.load(open(dict_path, encoding="utf-8")) if os.path.exists(dict_path) else {}
    new = {}
    for n, p in paragraphs(prs):
        m, _, _ = marked(p)
        if m and needs(m, mode) and m not in new:
            new[m] = old.get(m, "")
    json.dump(new, open(dict_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(len(new), "strings,", sum(1 for v in new.values() if not v), "empty")


def build(src, dict_path, out):
    tr = json.load(open(dict_path, encoding="utf-8"))
    vals = [v for v in tr.values() if v]
    ko = sum(1 for v in vals if HANGUL.search(v))
    lang = "ko-KR" if ko * 2 > len(vals) else "en-US"   # majority: an English deck may still quote a Korean doc code
    prs = Presentation(src)
    done = 0
    for n, p in paragraphs(prs):
        m, info, toks = marked(p)
        v = (tr.get(f"{n}|{m}") or tr.get(m)) if m else None
        if v and v != m:
            rebuild(p, v, info[0], info[1], toks, lang)
            done += 1
    prs.save(out)
    print(out, "paragraphs replaced:", done)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    cmd = sys.argv[1]
    if cmd == "extract":
        extract(sys.argv[2], sys.argv[3], sys.argv[4])
    else:
        build(sys.argv[2], sys.argv[3], sys.argv[4])
