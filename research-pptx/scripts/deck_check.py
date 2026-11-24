# -*- coding: utf-8 -*-
"""Read-only content check of a research deck (proposal / report). Never writes.

  python -I deck_check.py <deck.pptx> [--max-chars 450] [--min-pt 10] [--text]

Prints one line per slide (number, layout, characters of text, pictures / tables / charts, title) and then
a list of things to look at:
  - leftover markers: [확인 필요 …], [출처 필요 …], TODO, lorem ipsum, template prompt text
  - empty placeholders, slides without a title, body slides whose title is repeated on another slide
  - slides with more text than --max-chars (reading decks can go higher; talks should go lower)
  - text smaller than --min-pt
  - captions: "그림 N" / "Figure N" / "Fig. N" and "표 N" / "Table N" must run 1,2,3… without gaps or repeats
  - citations: every [n] used in the body needs an entry [n] on a 참고문헌 / References slide, and the reverse
--text also dumps the full text of every slide (for proofreading and for comparing with the outline).
These are prompts for a human look, not verdicts: a flagged slide can be fine.
"""
import re
import sys

from pptx import Presentation

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
MARKERS = re.compile(r"\[확인 필요|\[출처 필요|\bTODO\b|lorem|ipsum|텍스트를 입력|제목을 입력|Click to add|\bx{3,}\b", re.I)
FIG = re.compile(r"^\s*[\[<(]?\s*(그림|Figure|Fig\.?)\s*(\d+)", re.I)
TAB = re.compile(r"^\s*[\[<(]?\s*(표|Table)\s*(\d+)", re.I)
CITE = re.compile(r"\[(\d+(?:\s*[-–,]\s*\d+)*)\]")
REF_TITLE = re.compile(r"참고\s*문헌|references|bibliography|인용\s*문헌", re.I)


def walk(shapes):
    for s in shapes:
        if s.shape_type == 6:                       # group
            yield from walk(s.shapes)
        else:
            yield s


def paragraphs(shape):
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            yield p
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    yield p


def para_text(p):
    return "".join(r.text for r in p.runs).strip()


def expand(spec):
    out = set()
    for part in re.split(r"\s*,\s*", spec):
        m = re.match(r"(\d+)\s*[-–]\s*(\d+)$", part)
        if m:
            out.update(range(int(m.group(1)), int(m.group(2)) + 1))
        elif part.isdigit():
            out.add(int(part))
    return out


def sequence_issues(label, found):
    """found: list of (number, slide). Report gaps, repeats and out-of-order captions."""
    issues, seen, expect = [], {}, 1
    for num, slide in found:
        if num in seen:
            issues.append(f"{label} {num}: 번호 중복 (슬라이드 {seen[num]}, {slide})")
            continue
        seen[num] = slide
        if num != expect:
            issues.append(f"{label} {num} (슬라이드 {slide}): {label} {expect}이(가) 올 차례")
        expect = num + 1
    return issues


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        sys.exit(__doc__)
    opt = sys.argv[1:]
    max_chars = int(opt[opt.index("--max-chars") + 1]) if "--max-chars" in opt else 450
    min_pt = float(opt[opt.index("--min-pt") + 1]) if "--min-pt" in opt else 10
    if "--max-chars" in opt:
        args.remove(str(max_chars))
    if "--min-pt" in opt:
        args = [a for a in args if a != opt[opt.index("--min-pt") + 1]]
    dump = "--text" in opt

    prs = Presentation(args[0])
    issues, figs, tabs, cited, listed, titles = [], [], [], {}, set(), {}
    has_refs = False
    print("no | layout | chars | pic/table/chart | title")
    for n, slide in enumerate(prs.slides, 1):
        title_shape = slide.shapes.title
        title = title_shape.text_frame.text.strip().replace("\n", " / ") if title_shape is not None else ""
        is_ref = bool(REF_TITLE.search(title))
        has_refs = has_refs or is_ref
        chars = pics = tables = charts = 0
        small = set()
        texts = []
        for s in walk(slide.shapes):
            tag = s._element.tag.split("}")[1]
            pics += tag == "pic"
            tables += bool(getattr(s, "has_table", False) and s.has_table)
            charts += bool(getattr(s, "has_chart", False) and s.has_chart)
            if s.is_placeholder and s.has_text_frame and not s.text_frame.text.strip() \
                    and s.placeholder_format.type not in (13, 15, 16):      # slide number, footer, date
                issues.append(f"슬라이드 {n}: 빈 자리 표시자 '{s.name}'")
            for p in paragraphs(s):
                t = para_text(p)
                if not t:
                    continue
                texts.append(t)
                if s != title_shape:
                    chars += len(t)
                if MARKERS.search(t):
                    issues.append(f"슬라이드 {n}: 남은 표시 → {t[:50]}")
                for pat, store in ((FIG, figs), (TAB, tabs)):
                    m = pat.match(t)
                    if m:
                        store.append((int(m.group(2)), n))
                if is_ref:
                    m = re.match(r"\s*\[(\d+)\]", t)
                    if m:
                        listed.add(int(m.group(1)))
                else:
                    for m in CITE.finditer(t):
                        for k in expand(m.group(1)):
                            cited.setdefault(k, n)
                for r in p.runs:
                    sz = r.font.size
                    if sz is not None and sz.pt < min_pt and r.text.strip():
                        small.add(sz.pt)
        print(f"{n:>3} | {slide.slide_layout.name} | {chars} | {pics}/{tables}/{charts} | {title[:60]}")
        if dump:
            for t in texts:
                print("      " + t)
        if not title:
            issues.append(f"슬라이드 {n}: 제목 없음")
        else:
            titles.setdefault(title, []).append(n)
        if chars > max_chars:
            issues.append(f"슬라이드 {n}: 글자 {chars}자 (기준 {max_chars}자) → 나누거나 줄일지 검토")
        if small:
            issues.append(f"슬라이드 {n}: 작은 글자 {sorted(small)}pt (기준 {min_pt:g}pt)")

    for t, ns in titles.items():
        if len(ns) > 1:
            issues.append(f"같은 제목 '{t[:40]}' → 슬라이드 {ns} (이어지는 장이면 정상)")
    issues += sequence_issues("그림", figs) + sequence_issues("표", tabs)
    if cited or listed:
        if not has_refs:
            issues.append(f"본문에 인용 번호 {sorted(cited)}가 있으나 참고문헌 슬라이드가 없음")
        else:
            for k in sorted(set(cited) - listed):
                issues.append(f"인용 [{k}] (슬라이드 {cited[k]}): 참고문헌 목록에 없음")
            for k in sorted(listed - set(cited)):
                issues.append(f"참고문헌 [{k}]: 본문에서 인용되지 않음")

    print(f"\n그림 {len(figs)}개, 표 {len(tabs)}개, 인용 번호 {len(cited)}개, 참고문헌 항목 {len(listed)}개")
    print(f"\n== 살펴볼 것 {len(issues)}건")
    for i in issues:
        print("- " + i)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
