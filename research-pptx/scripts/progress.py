# -*- coding: utf-8 -*-
"""Per-slide work state kept in a file, so "done" means recorded, not remembered (and a later session can resume).

  python -I progress.py init <storyline.md> <work/progress.json>
      one row per "###" of the storyline. Run again after the storyline changed: rows are matched by title,
      known states are kept, new slides start as 대기, slides no longer in the storyline are dropped.
  python -I progress.py set <work/progress.json> <slides> <state> [note]
      <slides>: "3", "3,5-7" or "all" (numbers are storyline order, 1-based).
  python -I progress.py stage <work/progress.json> <stage 0-8>
      record which stage of the workflow is active (see SKILL.md "전체 흐름").
  python -I progress.py show <work/progress.json>
      checklist of stages, counts per state, and the slides that are not finished.

States
  대기        not built yet
  자료필요    cannot be built until the user supplies something (say what in the note)
  만듦        built in the deck, not yet looked at in a render
  검수함      rendered and checked
  확인필요    built, but something on it needs the user's decision (note says what)
  막힘        could not be built (note says why); the deck is not complete while any slide is here
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import storyline_check as sc  # noqa: E402

STATES = ["대기", "자료필요", "만듦", "검수함", "확인필요", "막힘"]
STAGES = ["입력 확인", "템플릿 분석", "스토리라인 확정", "자료 준비", "샘플 슬라이드 승인",
          "전체 조립", "검수", "수정 반복", "파생판"]


def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def save(path, data):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)


def numbers(spec, n):
    if spec == "all":
        return list(range(1, n + 1))
    out = set()
    for part in spec.split(","):
        a, _, b = part.partition("-")
        out.update(range(int(a), int(b or a) + 1))
    return [i for i in sorted(out) if 1 <= i <= n]


def init(md, path):
    doc = sc.parse(md)
    old = load(path) if os.path.exists(path) else {"stage": 2, "slides": []}
    known = {s["title"]: s for s in old["slides"]}
    slides = []
    for s in doc["slides"]:
        prev = known.get(s["title"], {})
        slides.append({"title": s["title"], "chapter": s["chapter"],
                       "state": prev.get("state", "대기"), "note": prev.get("note", "")})
    dropped = [t for t in known if t not in {s["title"] for s in slides}]
    save(path, {"stage": old.get("stage", 2), "slides": slides})
    print(f"progress: {path}  ({len(slides)}장, 새로 {sum(1 for s in slides if s['title'] not in known)}장, 빠짐 {len(dropped)}장)")


def show(path):
    data = load(path)
    stage = data.get("stage", 0)
    print("단계")
    for i, name in enumerate(STAGES):
        mark = "✔" if i < stage else ("▶" if i == stage else " ")
        print(f"  {mark} {i}. {name}")
    slides = data["slides"]
    counts = {st: sum(1 for s in slides if s["state"] == st) for st in STATES}
    print("\n슬라이드 " + str(len(slides)) + "장: " + ", ".join(f"{k} {v}" for k, v in counts.items() if v))
    open_rows = [(i, s) for i, s in enumerate(slides, 1) if s["state"] != "검수함"]
    if open_rows:
        print("끝나지 않은 장")
        for i, s in open_rows:
            print(f"  {i:>2}. [{s['state']}] {s['title']}" + (f"  — {s['note']}" if s["note"] else ""))
    else:
        print("모든 장 검수함")


def main():
    a = sys.argv[1:]
    if len(a) == 3 and a[0] == "init":
        init(a[1], a[2])
    elif len(a) >= 4 and a[0] == "set":
        if a[3] not in STATES:
            sys.exit("state must be one of: " + ", ".join(STATES))
        data = load(a[1])
        picked = numbers(a[2], len(data["slides"]))
        for i in picked:
            data["slides"][i - 1]["state"] = a[3]
            data["slides"][i - 1]["note"] = " ".join(a[4:])
        save(a[1], data)
        print(f"{len(picked)}장 → {a[3]}")
    elif len(a) == 3 and a[0] == "stage":
        data = load(a[1])
        data["stage"] = int(a[2])
        save(a[1], data)
        print(f"stage → {a[2]}. {STAGES[int(a[2])]}")
    elif len(a) == 2 and a[0] == "show":
        show(a[1])
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
