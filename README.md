# research-pptx

연구 제안서와 연구 보고서 PPT를 기준 양식에 맞춰 만드는 에이전트 스킬입니다. Codex에서 쓰도록 만들었고,
`SKILL.md` 형식을 읽는 다른 에이전트에서도 쓸 수 있습니다.

- 양식(기관 지정 양식, 지난 발표 자료)을 읽어 그 규칙대로 만듭니다.
- 무엇을 어떤 순서로 말할지를 **뼈대 덱**(PPT)으로 먼저 보여 주고, PowerPoint에서 고쳐 확정합니다.
- 대표 한두 장을 **샘플**로 먼저 확인받은 뒤 전체를 만듭니다.
- 자료에 없는 수치·결과·참고문헌은 지어내지 않고 `[확인 필요]`로 남깁니다.
- 결과물은 PowerPoint에서 고칠 수 있는 일반 슬라이드입니다(이미지 슬라이드가 아닙니다).

## 필요한 것

- Windows + PowerPoint
- Python 3.11 이상: `python -m pip install python-pptx lxml Pillow`
- Codex

## 설치

`research-pptx` 폴더를 통째로 스킬 폴더에 복사합니다.

```
C:\Users\<사용자>\.codex\skills\research-pptx\
```

Codex를 다시 열고 "어떤 스킬을 쓸 수 있어?"라고 물어 `research-pptx`가 보이는지 확인합니다. 보이지 않으면
`C:\Users\<사용자>\.agents\skills\`로 옮깁니다. 자세한 내용은 [`research-pptx/INSTALL.md`](research-pptx/INSTALL.md).

## 사용

작업 폴더에 `template/`(양식), `source/`(연구 자료), `figures/`(그림)를 만들고 그 폴더에서 Codex를 엽니다.

> template 폴더의 양식으로 연구 제안 발표 자료를 만들 거야. source 폴더 자료를 보고 스토리라인부터 잡아 줘.
> 발표는 15분이야.

사용 방법은 [`docs/research-pptx 사용 매뉴얼.pptx`](docs/)에 그림과 함께 정리돼 있습니다.

## 구성

```
research-pptx/
  SKILL.md            전체 흐름과 원칙
  INSTALL.md          설치와 처음 쓰는 법
  references/         단계별 상세 지침 (양식 분석, 스토리라인, 그림·참고문헌, 조립, 검수, 파생판)
  scripts/            뼈대 덱 만들기·읽기, 내용 점검, 렌더링, 진행 상태 기록 등
  assets/             스토리라인 빈 양식
docs/
  research-pptx 사용 매뉴얼.pptx
  build_manual.py     매뉴얼을 다시 만드는 스크립트
```

## 알아 둘 것

- 예시에 나오는 과제와 수치는 모두 가상입니다.
- 렌더링 스크립트(`ppt_export.ps1`)는 PowerPoint가 설치된 Windows에서만 동작합니다.
