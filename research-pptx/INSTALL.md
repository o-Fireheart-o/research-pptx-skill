# 설치와 사용 (Codex 기준)

이 파일은 사람이 읽는 안내다. 스킬 본문은 `SKILL.md`다.

## 준비물

- Windows + PowerPoint
- Python 3.11 이상과 패키지 세 개
  ```
  python -m pip install python-pptx lxml Pillow
  ```
- Codex (CLI 또는 앱)

## 설치

스킬은 GitHub 저장소 https://github.com/o-Fireheart-o/research-pptx-skill 에 있다. 아래 셋 중 편한
방법을 쓴다.

**1. Codex에게 시키기 (가장 쉬움)**

Codex에 이렇게 말한다.

> https://github.com/o-Fireheart-o/research-pptx-skill 저장소의 research-pptx 폴더와 research-pptx-init
> 폴더를 내 Codex 스킬 폴더에 설치해 줘.

**2. ZIP으로 받기**

1. 저장소 페이지에서 초록색 **Code** 버튼 > **Download ZIP**
2. 압축을 풀고, 안에 있는 `research-pptx` 폴더와 `research-pptx-init` 폴더를
   `C:\Users\<사용자>\.codex\skills\`에 복사

**3. git 명령으로 받기 (PowerShell)**

```
git clone https://github.com/o-Fireheart-o/research-pptx-skill.git
Copy-Item -Recurse research-pptx-skill\research-pptx "$env:USERPROFILE\.codex\skills\"
Copy-Item -Recurse research-pptx-skill\research-pptx-init "$env:USERPROFILE\.codex\skills\"
```

설치 뒤의 모습은 `C:\Users\<사용자>\.codex\skills\research-pptx\SKILL.md`다. Codex를 다시 열고 "어떤 스킬을
쓸 수 있어?"라고 물어 `research-pptx`가 보이는지 확인한다.

- 보이지 않으면 폴더를 `C:\Users\<사용자>\.agents\skills\`로 옮긴다. Codex 문서가 두 위치를 섞어 적고
  있다(2026-11 확인).
- 한 작업 폴더에서만 쓰려면 `<작업 폴더>\.agents\skills\research-pptx\`에 넣는다.
- 새 버전으로 바꿀 때는 같은 방법으로 다시 받아 두 폴더를 덮어쓴다.
- 스킬은 두 개다. `research-pptx`는 PPT를 만들고, `research-pptx-init`은 작업 폴더를 준비하고 다음에 할 일을
  알려 준다. 둘은 같은 스킬 폴더 안에 나란히 있어야 한다.

## 처음 쓸 때

1. 빈 폴더를 하나 만들고(또는 자료가 든 폴더에서) Codex를 열어 말한다.
   > 연구 PPT 폴더 만들어 줘.

   폴더가 자동으로 만들어지고, 무엇을 어디에 넣을지 대화창에서 알려 준다. 폴더에 이미 파일이 있으면
   어디로 옮길지 물어본다. (Codex에 원래 있는 `/init` 명령은 다른 기능이므로 위처럼 말로 한다.
   `$research-pptx-init`이라고 입력해도 된다.)
   ```
   내과제/
     template/   발표 양식이나 지난 발표 자료(.pptx)   ← 넣는다
     source/     공고문, 초안, 논문, 데이터, 지난 제안서·보고서   ← 넣는다
     figures/    쓸 그림   ← 있으면 넣는다
     output/, work/, 시작하기.md, AGENTS.md   ← 자동으로 만들어진다
   ```
   중간에 무엇을 해야 할지 모르겠으면 "지금 뭘 하면 돼?"라고 묻는다.
2. 그 폴더에서 Codex를 열고 말한다. 예:
   - "template에 있는 양식으로 연구 제안 발표 자료를 만들 거야. source 자료를 보고 스토리라인부터
     잡아 줘. 발표는 15분이야."
   - "지난 연차 보고(template)와 같은 형식으로 2차년도 보고 자료를 만들어 줘."
3. `storyboard.pptx`(뼈대 덱)가 만들어지면 PowerPoint로 연다. 1번 장이 전체 흐름도이고, 그 뒤로 실제
   슬라이드 한 장마다 기획 장이 하나씩 있다. 제목만 이어 읽어 이야기가 이어지는지 본다.
4. 고치고 싶으면 PowerPoint에서 바로 고친다. 순서는 보기 > 여러 슬라이드에서 끌어서 바꾸고, 제목과
   내용은 고쳐 쓰고, 필요 없는 장은 지운다. 만드는 쪽에 전할 말은 슬라이드 노트에 적는다. **저장하고
   PowerPoint를 닫은 뒤** "뼈대 덱 고쳤어, 읽어 줘"라고 한다. 말로 해도 된다("3번과 4번 순서 바꿔 줘").
5. 스토리라인이 정해지면 "이대로 만들어 줘"라고 한다. 먼저 **샘플 슬라이드 한두 장**
   (`output/sample.pptx`)을 보여 준다. 글의 양, 글자 크기, 배치가 마음에 드는지 보고 답하면 그 방식으로
   전체를 만든다. 결과는 `output/`에 생긴다.
6. 완성본을 PowerPoint에서 직접 고친 뒤 다시 맡길 때는 **저장하고 PowerPoint를 닫은 다음** 말한다
   ("내가 고친 파일 기준으로 7번 장 표만 바꿔 줘").

## 스토리라인을 글로 쓰고 싶을 때

뼈대 덱 대신 글 파일로 써도 된다. `assets/storyline_template.md`를 작업 폴더에 `storyline.md`로 복사해
쓴다. 제목만 적은 PPT 몇 장을 직접 만들어 넘겨도 같은 방식으로 읽는다. 글 파일의 규칙은 세 가지다.

- `## `으로 시작하면 장, `### `으로 시작하면 슬라이드 한 장
- `### ` 뒤에는 그 장에서 말하려는 결론을 한 문장으로
- 모르는 것은 `? `로 시작하는 줄에 적어 둔다

`###` 줄만 순서대로 써서 넘겨도 된다. 나머지는 자료를 보고 채워서 다시 보여 준다.

## 사용량을 아끼려면

- 확인을 두 번 요청한다(스토리라인, 샘플 슬라이드). 귀찮아도 여기서 답해 주는 편이 다 만든 뒤 고치는
  것보다 훨씬 적게 든다.
- 단계가 끝나면 **새 대화를 열어** 이어 가도 된다. "work/NOTES.md와 work/progress.json 읽고 이어서 해 줘"
  라고 하면 된다. 대화가 길어질수록 한 번 답할 때 드는 양이 커진다. 스토리라인을 확정한 뒤와 샘플을
  승인한 뒤가 끊기 좋은 때다.
- 고칠 것은 모아서 한 번에 말한다("7번 표 글자 키우고, 12번 그림 바꾸고, 15번 삭제").
- 같은 기관 양식을 다시 쓸 때는 지난번 분석 기록이 자동으로 쓰인다(문서를 마칠 때 "양식 기록을
  저장할까요?"에 그렇다고 답해 두면 된다). 기록은 `C:\Users\<사용자>\.research-pptx\profiles\`에 쌓인다.

## 알아 둘 것

- "완성했다"는 말은 정해 둔 기준을 다 채웠을 때만 한다. 자료가 없거나 막힌 장이 있으면 몇 번 장이 왜
  남았는지 알려 준다. 그럴듯한 값으로 채워 넣지 않는다.

- 자료에 없는 수치나 참고문헌은 채워 넣지 않고 `[확인 필요]`로 남긴다. 결과 보고의 "확인이 필요한
  것"을 꼭 본다.
- 양식 원본(`template/`)은 고치지 않는다. 고치기 전 파일은 `output/backup/`에 날짜와 함께 남는다.
- 작업 중 정한 규칙은 `work/NOTES.md`에 쌓인다. 며칠 뒤 이어서 할 때 "work/NOTES.md 읽고 이어서
  해 줘"라고 하면 된다.
