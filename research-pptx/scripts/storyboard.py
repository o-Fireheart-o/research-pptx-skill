# -*- coding: utf-8 -*-
"""Storyboard deck: the storyline as a PowerPoint the user reads and edits (one planning slide per real slide).

  python -I storyboard.py make <storyline.md> <out.pptx> [--template <template.pptx>] [--layout "<layout name>"]
      storyline.md -> storyboard deck. Slide 1 is a one-page flow map (chapters as columns, one box per slide);
      then one slide per "###": title = the slide's conclusion, body = content and "?" lines (red),
      dashed boxes = figures / tables wanted, speaker notes = 메모, small label top right = chapter.
      With --template the deck uses that file's slide size, theme and one of its layouts (--layout, default:
      the layout most slides use), so it already looks like the final deck. Without it: plain white 16:9.
  python -I storyboard.py read <storyboard.pptx> <storyline.md>
      storyboard deck (after the user reordered / rewrote / added / deleted slides) -> storyline.md.
      An existing storyline.md is kept as storyline.md.bak. Prints what changed in the slide titles.
  python -I storyboard.py map <storyboard.pptx>
      redraw the flow map on slide 1 from the current slides (after the user changed the deck).

What the user may do in PowerPoint, and how it is read back
  - drag slides to reorder (slide sorter view), delete slides, duplicate a slide and rewrite it
  - title            -> "### …"            body lines -> "- …" (a line starting with ? -> open question)
  - lines starting 그림: / 표: / 출처: anywhere on the slide, and the dashed boxes -> figures / tables / sources
  - speaker notes    -> 메모               chapter label (top right) -> "## …" (missing: same as slide before)
  - pictures pasted on a slide are reported in 메모 (they are not copied anywhere)
The flow map (slide 1) is a view: edits to its slide boxes are not read back. Its text box with "키: 값" lines
(종류, 발표, 핵심 주장 …) and its title ARE read back.
"""
import os
import re
import shutil
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import storyline_check as sc  # noqa: E402

P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
NAVY, GRAY, RED, LIGHT = RGBColor(0x1F, 0x3A, 0x5F), RGBColor(0x80, 0x80, 0x80), RGBColor(0xC0, 0x00, 0x00), RGBColor(0xEE, 0xF2, 0xF7)
BULLET = re.compile(r"^\s*[•\-–·▪➢□○]\s*")
KEYLINE = re.compile(r"^\s*(그림|표|출처|메모|시간|figure|fig|table|source|note|time)\s*[:：]\s*(.*)$", re.I)


# ---------------------------------------------------------------- storyline.md writer
def groups(doc):
    """[(chapter, [slides])] in slide order; chapters without slides are kept at their place in doc['chapters']."""
    out = []
    for s in doc["slides"]:
        if not out or out[-1][0] != s["chapter"]:
            out.append((s["chapter"], []))
        out[-1][1].append(s)
    names = [g[0] for g in out]
    for i, c in enumerate(doc["chapters"]):
        if c in names:
            continue
        before = [len(names) - 1 - names[::-1].index(x) for x in doc["chapters"][:i] if x in names]
        at = max(before) + 1 if before else 0
        out.insert(at, (c, []))
        names.insert(at, c)
    return out


def slide_md(s):
    out = ["", f"### {s['title']}"]
    out += ["  " * b["level"] + f"- {b['text']}" for b in s["bullets"]]
    out += s["text"]
    for key, field in (("그림", "figures"), ("표", "tables")):
        out += [f"{key}: {i['what']}" + (f" ({i['file']})" if i["file"] else "") for i in s[field]]
    out += [f"출처: {x}" for x in s["sources"]]
    out += [f"메모: {x}" for x in s["notes"]]
    if s["minutes"] is not None:
        out.append(f"시간: {s['minutes']:g}")
    out += [f"? {q}" for q in s["questions"]]
    return out


def dump_md(doc):
    out = [f"# {doc['title']}".rstrip(), ""]
    out += [f"{k}: {v}" for k, v in doc["meta"].items()]
    out += [f"? {q}" for q in doc["questions"]]
    for chapter, slides in groups(doc):
        out += ["", f"## {chapter}" if chapter else "## (장 없음)"]
        for s in slides:
            out += slide_md(s)
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------- helpers
def textbox(slide, name, x, y, w, h, size, bold=False, color=None, align=None, wrap=True):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tb.name = name
    tb.text_frame.word_wrap = wrap
    return tb, {"size": size, "bold": bold, "color": color, "align": align}


def write(tf, lines, st):
    """lines: list of (text, color or None). First line reuses the frame's first paragraph."""
    for i, (text, color) in enumerate(lines or [("", None)]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if st.get("align"):
            p.alignment = st["align"]
        r = p.add_run()
        r.text = text
        r.font.size = Pt(st["size"])
        r.font.bold = st.get("bold", False)
        c = color or st.get("color")
        if c is not None:
            r.font.color.rgb = c


def drop_all_slides(prs):
    lst = prs.slides._sldIdLst
    for sld in list(lst):
        prs.part.drop_rel(sld.rId)
        lst.remove(sld)


def pick_layout(prs, name, used):
    if name:
        for layout in prs.slide_layouts:
            if layout.name == name:
                return layout
        sys.exit(f"layout '{name}' not found. Layouts: " + ", ".join(l.name for l in prs.slide_layouts))
    if used:
        best = max(used, key=used.get)
        for layout in prs.slide_layouts:
            if layout.name == best:
                return layout
    for layout in prs.slide_layouts:
        if any(ph.placeholder_format.type in (1, 3) for ph in layout.placeholders):   # TITLE, CENTER_TITLE
            return layout
    return prs.slide_layouts[len(prs.slide_layouts) - 1]


def clear_placeholders(slide, keep_title=True):
    for ph in list(slide.placeholders):
        if keep_title and ph.placeholder_format.type in (1, 3):
            continue
        ph._element.getparent().remove(ph._element)


def is_map(slide):
    return any(s.name == "SB_meta" for s in slide.shapes)


# ---------------------------------------------------------------- flow map
def draw_map(prs, slide, doc):
    for s in list(slide.shapes):
        if s.name.startswith("SB_"):
            s._element.getparent().remove(s._element)
    W, H = prs.slide_width, prs.slide_height
    m = int(W * 0.03)
    tb, st = textbox(slide, "SB_maptitle", m, int(H * 0.03), W - 2 * m, int(H * 0.08), 22, True, NAVY)
    write(tb.text_frame, [(doc["title"] or "스토리라인", None)], st)
    meta = [(f"{k}: {v}", None) for k, v in doc["meta"].items()] or [("핵심 주장: ", None)]
    tb, st = textbox(slide, "SB_meta", m, int(H * 0.105), W - 2 * m, int(H * 0.16), 10, False, RGBColor(0x40, 0x40, 0x40))
    write(tb.text_frame, meta, st)

    chapters = groups(doc)
    if not chapters:
        return
    top, bottom = int(H * 0.31), int(H * 0.97)
    gap = int(W * 0.012)
    col_w = (W - 2 * m - gap * (len(chapters) - 1)) // len(chapters)
    head_h = int(H * 0.07)
    most = max(len(c[1]) for c in chapters)
    box_h = min(int(H * 0.085), (bottom - top - head_h - gap) // max(most, 1) - gap // 2)
    size = 11 if box_h > H * 0.07 and len(chapters) <= 5 else 9
    n, prev = 0, None
    for ci, (name, slides) in enumerate(chapters):
        x = m + ci * (col_w + gap)
        head = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(top), Emu(col_w), Emu(head_h))
        head.name = f"SB_chapter_{ci + 1}"
        head.fill.solid()
        head.fill.fore_color.rgb = NAVY if slides else GRAY          # gray header = chapter with no slide yet
        head.line.fill.background()
        head.text_frame.word_wrap = True
        write(head.text_frame, [(f"{name or '(장 없음)'}  ·  {len(slides)}장", None)],
              {"size": size + 1, "bold": True, "color": RGBColor(0xFF, 0xFF, 0xFF), "align": PP_ALIGN.CENTER})
        if prev is not None:
            cx = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, 0, 0)
            cx.name = f"SB_arrow_{ci}"
            cx.begin_connect(prev, 3)
            cx.end_connect(head, 1)
            cx.line.color.rgb = NAVY
        prev = head
        for si, s in enumerate(slides):
            n += 1
            y = top + head_h + gap // 2 + si * (box_h + gap // 2)
            box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(col_w), Emu(box_h))
            box.name = f"SB_slide_{n}"
            box.fill.solid()
            box.fill.fore_color.rgb = LIGHT
            box.line.color.rgb = RED if s["questions"] else RGBColor(0xB0, 0xB8, 0xC4)
            tf = box.text_frame
            tf.word_wrap = True
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf.margin_left = tf.margin_right = Pt(4)
            tf.margin_top = tf.margin_bottom = Pt(1)
            mark = "  ?" if s["questions"] else ""
            write(tf, [(f"{n}. {s['title']}{mark}", None)], {"size": size, "color": RGBColor(0x20, 0x20, 0x20)})


# ---------------------------------------------------------------- make
def make(md, out, template=None, layout_name=None):
    doc = sc.parse(md)
    if template:
        prs = Presentation(template)
        used = {}
        for s in list(prs.slides)[1:]:
            used[s.slide_layout.name] = used.get(s.slide_layout.name, 0) + 1
        layout = pick_layout(prs, layout_name, used)
        drop_all_slides(prs)
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
        layout = prs.slide_layouts[6]                      # Blank
    W, H = prs.slide_width, prs.slide_height

    cover = prs.slides.add_slide(layout)
    clear_placeholders(cover, keep_title=False)
    draw_map(prs, cover, doc)

    for s in doc["slides"]:
        slide = prs.slides.add_slide(layout)
        clear_placeholders(slide)
        title = slide.shapes.title
        if title is not None:
            title.name = "SB_title"
            title.text_frame.text = s["title"]
        else:
            tb, st = textbox(slide, "SB_title", int(W * 0.05), int(H * 0.05), int(W * 0.72), int(H * 0.11), 22, True, NAVY)
            write(tb.text_frame, [(s["title"], None)], st)
        tb, st = textbox(slide, "SB_chapter", int(W * 0.72), int(H * 0.015), int(W * 0.27), int(H * 0.05), 11, False, GRAY, PP_ALIGN.RIGHT)
        write(tb.text_frame, [(s["chapter"] or "", None)], st)

        visuals = [("그림", i) for i in s["figures"]] + [("표", i) for i in s["tables"]]
        body_w = int(W * 0.50) if visuals else int(W * 0.90)
        lines = [(("• " if b["level"] == 0 else "    – ") + b["text"], None) for b in s["bullets"]]
        lines += [(t, None) for t in s["text"]]
        lines += [(f"출처: {x}", GRAY) for x in s["sources"]]
        lines += [(f"? {q}", RED) for q in s["questions"]]
        tb, st = textbox(slide, "SB_body", int(W * 0.05), int(H * 0.22), body_w, int(H * 0.68), 16)
        write(tb.text_frame, lines or [("(내용을 적어 주세요)", GRAY)], st)

        if visuals:
            x, y0 = int(W * 0.58), int(H * 0.22)
            h = min(int(H * 0.30), int(H * 0.68) // len(visuals) - int(H * 0.02))
            for i, (kind, item) in enumerate(visuals):
                box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y0 + i * (h + int(H * 0.02))), Emu(int(W * 0.37)), Emu(h))
                box.name = f"SB_fig_{i + 1}"
                box.fill.solid()
                box.fill.fore_color.rgb = RGBColor(0xF7, 0xF7, 0xF7)
                box.line.color.rgb = GRAY
                box.line.dash_style = MSO_LINE.DASH
                box.text_frame.word_wrap = True
                text = [(f"{kind}: {item['what']}", None)]
                if item["file"]:
                    text.append((f"({item['file']})", GRAY))
                write(box.text_frame, text, {"size": 13, "color": RGBColor(0x40, 0x40, 0x40), "align": PP_ALIGN.CENTER})
        memo = list(s["notes"]) + ([f"시간: {s['minutes']:g}"] if s["minutes"] is not None else [])
        if memo:
            slide.notes_slide.notes_text_frame.text = "\n".join(memo)
    prs.save(out)
    print(f"storyboard: {out}  (흐름도 1장 + 슬라이드 {len(doc['slides'])}장, layout '{layout.name}')")


# ---------------------------------------------------------------- read
def shape_lines(shape):
    for p in shape.text_frame.paragraphs:
        t = "".join(r.text for r in p.runs).replace("\v", " ").rstrip()
        if t.strip():
            yield p.level, t


def read_deck(path):
    prs = Presentation(path)
    doc = {"title": "", "meta": {}, "chapters": [], "slides": [], "questions": []}
    chapter = None
    for no, slide in enumerate(prs.slides, 1):
        shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        if is_map(slide):
            for s in shapes:
                if s.name == "SB_maptitle":
                    doc["title"] = s.text_frame.text.strip()
                elif s.name.startswith("SB_chapter_"):
                    name = re.sub(r"\s*·\s*\d+장\s*$", "", s.text_frame.text.strip())
                    if name and name != "(장 없음)":
                        doc["chapters"].append(name)
                elif s.name == "SB_meta":
                    for _, t in shape_lines(s):
                        if t.strip().startswith(("?", "？")):
                            doc["questions"].append(t.strip()[1:].strip())
                        elif re.match(r"^[^:：]{1,12}[:：]", t.strip()):
                            k, v = re.split(r"[:：]", t.strip(), 1)
                            doc["meta"][k.strip()] = v.strip()
            continue
        title_shape = next((s for s in shapes if s.name == "SB_title"), None) or \
            (slide.shapes.title if slide.shapes.title is not None and slide.shapes.title.text_frame.text.strip() else None)
        rest = [s for s in shapes if s is not title_shape and s.name != "SB_chapter"]
        if title_shape is None and rest:                     # slide the user built from scratch: topmost text is the title
            title_shape = min(rest, key=lambda s: (s.top or 0, s.left or 0))
            rest.remove(title_shape)
        label = next((s for s in shapes if s.name == "SB_chapter"), None)
        if label is not None:
            chapter = label.text_frame.text.strip() or chapter
        if title_shape is None and not rest:
            continue                                          # empty slide
        title = " ".join(title_shape.text_frame.text.split()) if title_shape is not None else f"(제목 없음, 뼈대 덱 {no}번)"
        s = sc.new_slide(title, chapter, no)
        if chapter not in doc["chapters"]:
            doc["chapters"].append(chapter)
        for shp in sorted(rest, key=lambda x: (x.top or 0, x.left or 0)):
            pending = None
            for level, t in shape_lines(shp):
                st = t.strip()
                if st in ("(내용을 적어 주세요)",):
                    continue
                m = KEYLINE.match(st)
                fm = re.match(r"^\(([^()]+\.[A-Za-z0-9]{2,5})\)$", st)
                if fm and pending is not None:               # "(file)" line under a 그림:/표: line
                    pending["file"] = fm.group(1)
                elif st.startswith(("?", "？")):
                    s["questions"].append(st[1:].strip())
                elif m:
                    field = sc.KEYS[m.group(1).lower()]
                    value = m.group(2).strip()
                    if field in ("figures", "tables"):
                        f2 = sc.FILE_RE.search(value)
                        pending = {"what": sc.FILE_RE.sub("", value).strip(), "file": f2.group(1) if f2 else None}
                        s[field].append(pending)
                    elif field == "minutes":
                        try:
                            s["minutes"] = float(re.sub(r"[^\d.]", "", value))
                        except ValueError:
                            pass
                    else:
                        s[field].append(value)
                else:
                    sub = level > 0 or bool(re.match(r"^\s{2,}", t)) or st.startswith("–")
                    s["bullets"].append({"level": 1 if sub else 0, "text": BULLET.sub("", st)})
        if slide.has_notes_slide:
            for line in slide.notes_slide.notes_text_frame.text.splitlines():
                m = KEYLINE.match(line.strip())
                if m and sc.KEYS[m.group(1).lower()] == "minutes":
                    try:
                        s["minutes"] = float(re.sub(r"[^\d.]", "", m.group(2)))
                    except ValueError:
                        pass
                elif line.strip():
                    s["notes"].append(line.strip())
        pics = sum(1 for x in slide.shapes if x._element.tag == f"{P_NS}pic")
        if pics:
            s["notes"].append(f"뼈대 덱 {no}번 장에 그림 {pics}개를 직접 넣어 둠")
        doc["slides"].append(s)
    return prs, doc


def read(deck, md):
    _, doc = read_deck(deck)
    old = [s["title"] for s in sc.parse(md)["slides"]] if os.path.exists(md) else []
    if os.path.exists(md):
        shutil.copyfile(md, md + ".bak")
    with open(md, "w", encoding="utf-8") as f:
        f.write(dump_md(doc))
    new = [s["title"] for s in doc["slides"]]
    runs = [g[0] for g in groups(doc) if g[1]]
    for c in sorted({c for c in runs if runs.count(c) > 1}, key=str):
        print(f"  ! 장 '{c}'의 슬라이드가 떨어져 있음 → 옮긴 장의 오른쪽 위 장 이름이 예전 것인지 확인")
    print(f"storyline: {md}  (슬라이드 {len(new)}장)" + ("  이전 파일은 .bak으로 보관" if old else ""))
    if old:
        added, removed = [t for t in new if t not in old], [t for t in old if t not in new]
        kept_old, kept_new = [t for t in old if t in new], [t for t in new if t in old]
        for t in added:
            print(f"  + 새로 생기거나 제목이 바뀜: {t}")
        for t in removed:
            print(f"  - 없어지거나 제목이 바뀜: {t}")
        if kept_old != kept_new:
            print("  ~ 순서가 바뀜")
        if not added and not removed and kept_old == kept_new:
            print("  제목과 순서는 그대로 (내용만 바뀌었을 수 있음)")


def remap(deck):
    prs, doc = read_deck(deck)
    target = next((s for s in prs.slides if is_map(s)), None)
    if target is None:
        sys.exit("흐름도 슬라이드(SB_meta 상자가 있는 장)를 찾지 못함. make로 다시 만들 것.")
    draw_map(prs, target, doc)
    prs.save(deck)
    print(f"flow map redrawn: {deck}  (슬라이드 {len(doc['slides'])}장)")


def main():
    a = sys.argv[1:]
    opt = {}
    for key in ("--template", "--layout"):
        if key in a:
            i = a.index(key)
            opt[key] = a[i + 1]
            del a[i:i + 2]
    if len(a) == 3 and a[0] == "make":
        make(a[1], a[2], opt.get("--template"), opt.get("--layout"))
    elif len(a) == 3 and a[0] == "read":
        read(a[1], a[2])
    elif len(a) == 2 and a[0] == "map":
        remap(a[1])
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
