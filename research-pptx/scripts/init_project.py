# -*- coding: utf-8 -*-
"""Set up (or check) a working folder for one research deck, and say what to do next.

  python -I init_project.py [<folder>] [--name "<문서 이름>"] [--sort]

  <folder>   working folder; default: the current folder. Created if missing.
  --name     document / project name written into the notes (optional).
  --sort     move loose files that sit directly in the folder into source/ and figures/ by type
             (documents and data -> source, images -> figures). .pptx / .potx files are never moved
             automatically: they may be the template or source material, so they are only listed.

Creates what is missing and never overwrites or deletes anything:
  template/  source/  figures/  output/backup/  work/
  work/NOTES.md      running notes (rules decided, which file is the master, what is still missing)
  AGENTS.md          tells the agent to use the research-pptx skill and to read work/NOTES.md first
  시작하기.md         what goes where and what to say (for the user)
Safe to run again at any time: it then only reports the state of the folder and the next step.
The printed report is written for the agent to pass on to the user in plain words.
"""
import datetime
import os
import shutil
import sys

DOCS = {".pdf", ".doc", ".docx", ".hwp", ".hwpx", ".txt", ".md", ".xls", ".xlsx", ".csv", ".tsv", ".json", ".bib", ".ris"}
IMAGES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".svg", ".emf", ".wmf", ".webp"}
DECKS = {".pptx", ".potx", ".ppt"}
OWN = {"AGENTS.md", "시작하기.md", "storyline.md", "storyboard.pptx"}

AGENTS = """# 작업 안내

이 폴더는 연구 제안서·연구 보고서 PPT 한 건을 만드는 작업 폴더다.

- `research-pptx` 스킬을 쓴다.
- 작업을 시작하거나 이어 갈 때 `work/NOTES.md`를 먼저 읽고, `work/progress.json`이 있으면 그것도 읽는다.
- `template/`의 양식 원본은 고치지 않는다. 완성본은 `output/`에 만든다.
- 사용자가 정한 규칙과 파일의 현재 상태는 `work/NOTES.md`에 날짜와 함께 적는다.
"""

NOTES = """# 작업 기록

- 문서: {name}
- 폴더를 만든 날: {today}

## 정해진 것
(사용자가 정한 규칙, 용어, 샘플 슬라이드에서 승인된 값을 날짜와 함께 적는다)

## 파일 상태
(어느 파일이 원본인지, 다시 만들면 안 되는 파일)

## 아직 받지 못한 자료
"""

GUIDE = """# 시작하기

이 폴더는 연구 발표 자료 한 건을 만드는 작업 폴더입니다.

## 1. 자료를 넣습니다

| 폴더 | 넣을 것 |
|---|---|
| `template` | 발표 양식이나 지난 발표 자료 (.pptx). **꼭 필요합니다** |
| `source` | 공고문, 초안, 논문, 데이터, 지난 제안서·보고서. **하나 이상 필요합니다** |
| `figures` | 쓸 그림과 그래프 (없어도 됩니다) |

## 2. 이렇게 말합니다

> template 폴더의 양식으로 연구 제안 발표 자료를 만들 거야. source 폴더 자료를 보고 스토리라인부터 잡아 줘.
> 발표는 15분이야.

문서 종류(제안 발표 / 연차 보고 / 최종 보고 / 회의 자료), 발표 시간, 분량 제한을 함께 말해 주면 좋습니다.

## 3. 그다음

1. `storyboard.pptx`(뼈대 덱)가 만들어집니다. PowerPoint에서 열어 흐름을 보고 고칩니다.
   저장하고 닫은 뒤 "뼈대 덱 고쳤어, 읽어 줘"라고 합니다.
2. "이대로 만들어 줘"라고 하면 샘플 슬라이드 한두 장을 먼저 보여 줍니다.
3. 샘플이 마음에 들면 전체를 만듭니다. 결과는 `output` 폴더에 생깁니다.

## 나머지 폴더

- `output`: 샘플과 완성본. `output/backup`에는 고치기 전 파일이 날짜와 함께 남습니다.
- `work`: 작업 기록. 열어 보지 않아도 됩니다. 며칠 뒤 이어서 할 때 "work/NOTES.md 읽고 이어서 해 줘"라고
  하면 됩니다.

지금 무엇을 해야 하는지 모르겠으면 "지금 뭘 하면 돼?"라고 물어보세요.
"""


def files_in(path):
    if not os.path.isdir(path):
        return []
    return sorted(f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f)) and not f.startswith(("~$", ".")))


def write_new(path, content, made):
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        made.append(os.path.basename(path))


def main():
    args = sys.argv[1:]
    name = ""
    if "--name" in args:
        i = args.index("--name")
        name = args[i + 1]
        del args[i:i + 2]
    sort = "--sort" in args
    args = [a for a in args if a != "--sort"]
    root = os.path.abspath(args[0] if args else ".")
    made = []
    for d in ("template", "source", "figures", os.path.join("output", "backup"), "work"):
        p = os.path.join(root, d)
        if not os.path.isdir(p):
            os.makedirs(p)
            made.append(d.replace(os.sep, "/") + "/")
    today = datetime.date.today().isoformat()
    write_new(os.path.join(root, "AGENTS.md"), AGENTS, made)
    write_new(os.path.join(root, "work", "NOTES.md"), NOTES.format(name=name or "(아직 정하지 않음)", today=today), made)
    write_new(os.path.join(root, "시작하기.md"), GUIDE, made)

    loose = [f for f in files_in(root) if f not in OWN]
    docs = [f for f in loose if os.path.splitext(f)[1].lower() in DOCS]
    imgs = [f for f in loose if os.path.splitext(f)[1].lower() in IMAGES]
    decks = [f for f in loose if os.path.splitext(f)[1].lower() in DECKS]
    moved = []
    if sort:
        for group, dest in ((docs, "source"), (imgs, "figures")):
            for f in group:
                target = os.path.join(root, dest, f)
                if os.path.exists(target):
                    continue                              # never overwrite
                shutil.move(os.path.join(root, f), target)
                moved.append(f"{f} → {dest}/")
        docs = [f for f in docs if os.path.exists(os.path.join(root, f))]
        imgs = [f for f in imgs if os.path.exists(os.path.join(root, f))]

    tpl = [f for f in files_in(os.path.join(root, "template")) if os.path.splitext(f)[1].lower() in DECKS]
    src = files_in(os.path.join(root, "source"))
    fig = files_in(os.path.join(root, "figures"))
    out = [f for f in files_in(os.path.join(root, "output")) if f.lower().endswith(".pptx")]
    has_story = os.path.exists(os.path.join(root, "storyline.md"))
    has_board = os.path.exists(os.path.join(root, "storyboard.pptx"))
    has_progress = os.path.exists(os.path.join(root, "work", "progress.json"))

    print(f"작업 폴더: {root}")
    print("새로 만든 것: " + (", ".join(made) if made else "없음 (이미 준비된 폴더)"))
    if moved:
        print("옮긴 파일: " + "; ".join(moved))
    print(f"\n상태: template 양식 {len(tpl)}개 / source 자료 {len(src)}개 / figures 그림 {len(fig)}개"
          + (f" / output 덱 {len(out)}개" if out else ""))
    if tpl:
        print("  양식: " + ", ".join(tpl))
    if decks or docs or imgs:
        print("\n폴더 바로 아래에 있는 파일 (아직 분류 안 됨)")
        if decks:
            print("  PPT (양식인지 참고 자료인지 사용자에게 물어볼 것): " + ", ".join(decks))
        if docs:
            print("  문서·데이터 (source로 옮길 후보): " + ", ".join(docs))
        if imgs:
            print("  그림 (figures로 옮길 후보): " + ", ".join(imgs))
        if not sort and (docs or imgs):
            print("  → 사용자가 좋다고 하면 --sort로 다시 실행해 옮긴다. PPT는 답을 듣고 직접 옮긴다.")

    print("\n다음에 할 일")
    step = 1
    if out or has_progress:
        print(f"  {step}. 진행 중인 작업이 있다. work/NOTES.md와 work/progress.json을 읽고 이어 간다.")
    elif has_story or has_board:
        print(f"  {step}. 스토리라인이 이미 있다. 뼈대 덱(storyboard.pptx)을 확인받고 확정한다.")
    else:
        if not tpl:
            print(f"  {step}. [사용자] template 폴더에 발표 양식이나 지난 발표 자료(.pptx)를 넣는다  ← 꼭 필요")
            step += 1
        if not src:
            print(f"  {step}. [사용자] source 폴더에 공고문, 초안, 논문, 데이터 등 연구 자료를 넣는다  ← 하나 이상 필요")
            step += 1
        if not fig:
            print(f"  {step}. [사용자] 쓸 그림이 있으면 figures 폴더에 넣는다 (없어도 됨)")
            step += 1
        ready = bool(tpl) and bool(src)
        print(f"  {step}. [사용자] " + ("준비됐다. 이렇게 말하면 시작한다:" if ready else "다 넣은 뒤 이렇게 말한다:"))
        print("       \"template 양식으로 연구 제안 발표 자료를 만들 거야. source 자료를 보고 스토리라인부터 잡아 줘. 발표는 15분이야.\"")
        print("       (문서 종류, 발표 시간, 분량 제한을 함께 말해 주면 좋다)")
    print("\n안내 파일: 시작하기.md (같은 내용이 적혀 있다)")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
