# -*- coding: utf-8 -*-
"""Read a storyline file (see references/storyline-format.md) and report what is in it. Never changes the file.

  python -I storyline_check.py <storyline.md> [--json <out.json>]

Format in one glance
  # 문서 제목                      document title
  키: 값                           settings before the first chapter (종류, 발표, 독자, 핵심 주장, 평가 …)
  ## 장 제목                       chapter
  ### 슬라이드의 결론 한 문장        one slide
  - 내용                           what goes on the slide (indent for sub-points)
  그림: / 표: 설명 (파일)           figure or table wanted; the file in parentheses if it exists
  출처: …                          source of the numbers / claims
  메모: …                          note to the writer (not shown on the slide)
  시간: 1.5                        talk minutes for this slide (optional)
  ? …                              open question, not decided yet

Prints a slide table, the share of each chapter, an estimated talk time and a list of things to settle.
--json writes the parsed storyline for the build script.
"""
import json
import re
import sys

KEYS = {"그림": "figures", "figure": "figures", "fig": "figures", "표": "tables", "table": "tables",
        "출처": "sources", "source": "sources", "메모": "notes", "note": "notes", "시간": "minutes", "time": "minutes"}
KEY_RE = re.compile(r"^\s*([A-Za-z가-힣 ]{1,12})\s*[:：]\s*(.*)$")
FILE_RE = re.compile(r"\(([^()]+\.[A-Za-z0-9]{2,5})\)\s*$")


def new_slide(title, chapter, line):
    return {"title": title, "chapter": chapter, "line": line, "bullets": [], "text": [], "figures": [],
            "tables": [], "sources": [], "notes": [], "questions": [], "minutes": None}


def parse(path):
    doc = {"title": "", "meta": {}, "chapters": [], "slides": [], "questions": []}
    chapter, slide, comment = None, None, False
    with open(path, encoding="utf-8-sig") as f:
        for no, raw in enumerate(f, 1):
            line = raw.rstrip()
            s = line.strip()
            if comment or s.startswith("<!--"):          # <!-- ... --> blocks (how-to text) are skipped
                comment = "-->" not in s
                continue
            if not s:
                continue
            if s == "###" or s.startswith("### "):
                title = s[3:].strip()
                if not title:                            # empty "###" left from the template
                    slide = None
                    continue
                slide = new_slide(title, chapter, no)
                doc["slides"].append(slide)
            elif s.startswith("## "):
                chapter, slide = s[3:].strip(), None
                doc["chapters"].append(chapter)
            elif s.startswith("# "):
                doc["title"] = s[2:].strip()
            elif s.startswith("?") or s.startswith("？"):
                q = s[1:].strip()
                (slide["questions"] if slide else doc["questions"]).append(q)
            elif re.match(r"^\s*[-*•]\s+", line):
                depth = (len(line) - len(line.lstrip())) // 2
                text = re.sub(r"^\s*[-*•]\s+", "", line)
                if slide:
                    slide["bullets"].append({"level": depth, "text": text})
            else:
                m = KEY_RE.match(s)
                key = m.group(1).strip().lower() if m else None
                if slide is None:
                    if m:
                        doc["meta"][m.group(1).strip()] = m.group(2).strip()
                elif key in KEYS:
                    field, value = KEYS[key], m.group(2).strip()
                    if field == "minutes":
                        try:
                            slide["minutes"] = float(re.sub(r"[^\d.]", "", value))
                        except ValueError:
                            pass
                    elif field in ("figures", "tables"):
                        fm = FILE_RE.search(value)
                        slide[field].append({"what": FILE_RE.sub("", value).strip(), "file": fm.group(1) if fm else None})
                    else:
                        slide[field].append(value)
                else:
                    slide["text"].append(s)
    return doc


def talk_minutes(meta):
    for k, v in meta.items():
        if k.startswith("발표") or k.lower().startswith("talk"):
            m = re.search(r"(\d+(?:\.\d+)?)\s*분", v) or re.search(r"(\d+(?:\.\d+)?)", v)
            if m:
                return float(m.group(1))
    return None


def report(doc):
    slides = doc["slides"]
    print(f"문서: {doc['title'] or '(제목 없음)'}")
    for k, v in doc["meta"].items():
        print(f"  {k}: {v}")
    print(f"\n슬라이드 {len(slides)}장 / 장 {len(doc['chapters'])}개\n")
    print("no | 장 | 제목 | 내용 | 그림/표 | 표시")
    issues = []
    for n, s in enumerate(slides, 1):
        flags = []
        body = len(s["bullets"]) + len(s["text"])
        if body == 0:
            flags.append("내용 없음")
            issues.append(f"{n}. {s['title']}: 내용이 비어 있음 → 자료에서 채우거나 직접 적기")
        if len(s["title"]) < 9:
            flags.append("제목이 주제어")
            issues.append(f"{n}. {s['title']}: 제목이 결론 문장이 아님 → '무엇이 어떻다'로")
        top = sum(1 for b in s["bullets"] if b["level"] == 0)
        if top > 5:
            flags.append(f"항목 {top}개")
            issues.append(f"{n}. {s['title']}: 큰 항목이 {top}개 → 두 장으로 나눌지 검토")
        for kind, items in (("그림", s["figures"]), ("표", s["tables"])):
            for it in items:
                if not it["file"]:
                    issues.append(f"{n}. {s['title']}: {kind} '{it['what']}' 파일 없음 → 자료 필요 또는 새로 그림")
        for q in s["questions"]:
            flags.append("?")
            issues.append(f"{n}. {s['title']}: ? {q}")
        vis = len(s["figures"]) + len(s["tables"])
        print(f"{n:>2} | {s['chapter'] or '-'} | {s['title']} | {body} | {vis} | {', '.join(flags)}")
    for q in doc["questions"]:
        issues.append(f"문서 전체: ? {q}")
    for c in doc["chapters"]:
        if not any(s["chapter"] == c for s in slides):
            issues.append(f"장 '{c}': 슬라이드 없음")
    if "핵심 주장" not in doc["meta"] and not any("주장" in k for k in doc["meta"]):
        issues.append("문서 전체: '핵심 주장:' 한 문장이 없음")

    if slides:
        print("\n장별 분량")
        for c in doc["chapters"]:
            k = sum(1 for s in slides if s["chapter"] == c)
            print(f"  {c}: {k}장 ({k * 100 // len(slides)}%)")
    total = talk_minutes(doc["meta"])
    given = sum(s["minutes"] or 0 for s in slides)
    if total and slides:
        rest = [s for s in slides if s["minutes"] is None]
        per = (total - given) / len(rest) if rest else 0
        print(f"\n발표 {total:g}분 → 장당 평균 {total / len(slides):.1f}분"
              + (f" (시간을 적지 않은 {len(rest)}장은 장당 {per:.1f}분)" if given and rest else ""))
        if total / len(slides) < 0.7:
            issues.append(f"문서 전체: 발표 {total:g}분에 {len(slides)}장 → 장당 1분이 안 됨. 줄이거나 보조 장표로")
    print(f"\n정할 것 {len(issues)}건")
    for i in issues:
        print("- " + i)


def main():
    args = sys.argv[1:]
    if not args:
        sys.exit(__doc__)
    doc = parse(args[0])
    report(doc)
    if "--json" in args:
        out = args[args.index("--json") + 1]
        with open(out, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False, indent=1)
        print(f"\njson → {out}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
